VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReportEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmReportEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8535
   StartUpPosition =   1  '����������
   Begin VB.Frame fraGroups 
      Caption         =   "����������"
      Height          =   3555
      Left            =   4200
      TabIndex        =   9
      Top             =   30
      Width           =   4215
      Begin VSFlex8Ctl.VSFlexGrid vsfGroups 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3975
         _cx             =   1989548899
         _cy             =   1989547418
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   12
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6000
      TabIndex        =   11
      Top             =   3720
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   4050
      Begin VB.ComboBox cboClass 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txt˵�� 
         BackColor       =   &H00FFFFFF&
         Height          =   1920
         Left            =   855
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1485
         Width           =   3000
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   855
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1065
         Width           =   3000
      End
      Begin VB.TextBox txt��� 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         MaxLength       =   20
         TabIndex        =   4
         Top             =   645
         Width           =   1500
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "˵��"
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   705
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmReportEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_GROUPS_COLS As String = _
    "ѡ��,,3,450,B|���,,3,1500|����,,3,1500|ID,,0,n"
    
Private mbytMode As Byte            '0-����1-�����飻2-�����ࣻ3-�������ӱ���
Private mlngSys As Long
Private mlngReportID As Long
Private mlngGroupID As Long
Private mlngClassID As Long
Private mstr���� As String
Private mstrOld���� As String
Private mstr���� As String
Private mstr˵�� As String
Private mstrOld˵�� As String
Private mblnOK As Boolean
Private mlngModule As Long

Private WithEvents mvsfGroups As clsVSFlexGridEx
Attribute mvsfGroups.VB_VarHelpID = -1

Public Function ShowMe(ByVal frmParent As Object, ByVal lngSys As Long _
    , ByVal bytMode As Byte, ByVal lngModule As Long _
    , Optional ByRef lngGroupID As Long, Optional ByRef lngReportID As Long _
    , Optional ByRef str���� As String, Optional ByRef str���� As String _
    , Optional ByRef str˵�� As String) As Boolean
    
    mlngSys = lngSys
    mbytMode = bytMode
    mlngModule = lngModule
    mlngReportID = lngReportID
    mlngGroupID = lngGroupID
    mstr���� = str����: mstrOld���� = str����
    mstr���� = str����
    mstr˵�� = str˵��: mstrOld˵�� = str˵��
    
    If bytMode = 2 Then
        mlngClassID = lngReportID
    ElseIf bytMode = 1 Then
        mlngClassID = GetClassID(lngGroupID, True)
    Else
        mlngClassID = GetClassID(lngReportID)
    End If
    
    Set mvsfGroups = New clsVSFlexGridEx
    
    Me.Show vbModal, frmParent
    str���� = mstr����
    str���� = mstr����
    str˵�� = mstr˵��
    ShowMe = mblnOK
End Function

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    If mbytMode = Val("2-������") Then
        If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsCheck As New ADODB.Recordset
    Dim strSQL As String, strOldCode As String, strOldName As String, strOld˵�� As String

    Dim intOrder As Integer
    Dim arrSQL() As Variant
    Dim i As Long, lngClassID As Long, lngTemp As Long, lngProgID As Long
    Dim blnTrans As Boolean
    
    If UCase(Me.ActiveControl.name) <> UCase("cmdOK") Then
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    End If
    
    arrSQL = Array()
    If Not CheckFormInput(Me) Then Exit Sub
    
    If mbytMode = Val("2-������") Then
        If Trim(txt����.Text) = "����" Then
            MsgBox "�����С�Ϊ���ౣ���ؼ��֣����޸ģ�", vbInformation, App.Title
            txt����.SetFocus: Exit Sub
        End If
    Else
        If Trim(txt���.Text) = "" Then
            MsgBox "�����뱨��" & IIF(mbytMode = 1, "��", "") & "�ı�ţ�", vbInformation, App.Title
            txt���.SetFocus: Exit Sub
        End If
    End If
    
    If Trim(txt����.Text) = "" Then
        Select Case mbytMode
        Case 2
            MsgBox "�����롰�����ࡱ�����ƣ�", vbInformation, App.Title
        Case 1
            MsgBox "�����롰�����顱�����ƣ�", vbInformation, App.Title
        Case Else
            MsgBox "�����롰���������ƣ�", vbInformation, App.Title
        End Select
        txt����.SetFocus
        Exit Sub
    Else
        txt����.Text = ConvertSBC(txt����.Text)
    End If
    
    If Not CheckLen(txt���, 20, "���") Then Exit Sub
    If Not CheckLen(txt����, 30, "����") Then Exit Sub
    If Not CheckLen(txt˵��, 255, "˵��") Then Exit Sub
    
    On Error GoTo hErr
    
    '���
    If mbytMode = Val("2-������") Then
        '������
        If mlngClassID = 0 Then
            strSQL = "Select ���� From zlRPTClasses Where Upper(����) = [1]"
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, UCase(txt����.Text))
            If rsCheck.RecordCount > 0 Then
                MsgBox "�������ࡰ" & txt����.Text & "���ظ���", vbInformation, App.Title
                txt����.SetFocus
                Exit Sub
            End If
            rsCheck.Close
        End If
    Else
        '��Ų����ظ�(����������)
        If CheckExist("zlReports", "���", txt���.Text, mlngReportID) Then
            MsgBox "�ñ���Ѿ���ʹ��,���������룡", vbInformation, App.Title
            txt���.SetFocus: Exit Sub
        End If
        If CheckExist("zlRPTGroups", "���", txt���.Text, mlngGroupID) Then
            MsgBox "�ñ���Ѿ���ʹ��,���������룡", vbInformation, App.Title
            txt���.SetFocus: Exit Sub
        End If
        
        If mlngGroupID <> 0 And mbytMode <> Val("1-������") Then
            strSQL = _
                "Select 1 From zlRPTSubs A,zlReports B " & vbCrLf & _
                "Where B.����=[1] And A.����ID=B.ID And A.��ID=[2]" & _
                IIF(mlngReportID = 0, "", " And ����ID<>[3]")
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, txt����.Text, mlngGroupID, mlngReportID)
            If Not rsCheck.EOF Then
                MsgBox "�ñ��������Ѿ�������ͬ���Ƶı���", vbInformation, App.Title
                txt����.SetFocus: Exit Sub
            End If
        End If
    End If
    
    strOldCode = mstr����: strOldName = mstrOld����: strOld˵�� = mstrOld˵��
    mstr���� = txt����.Text: mstr���� = txt���.Text: mstr˵�� = txt˵��.Text
    If cboClass.ListIndex >= 0 Then
        lngClassID = cboClass.ItemData(cboClass.ListIndex)
    End If
    
    '����
    Select Case mbytMode
    Case Val("2-������")
        If mlngClassID = 0 Then
            '����
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlRPTClasses(ID,�ϼ�ID,����,˵��) " & vbCrLf & _
                "Values " & vbCrLf & _
                "(zlRPTClasses_ID.nextval " & vbCrLf & _
                "," & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                ",'" & mstr���� & "'" & vbCrLf & _
                "," & IIF(mstr˵�� = "", "Null", "'" & mstr˵�� & "'") & _
                ")"
        ElseIf Not (strOldName = mstr���� And strOld˵�� = mstr˵�� And lngClassID = mlngGroupID) Then
            '�޸�
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Update zlRPTClasses " & vbCrLf & _
                "Set �ϼ�ID = " & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                "   ,���� = '" & mstr���� & "'" & vbCrLf & _
                "   ,˵�� = " & IIF(mstr˵�� = "", "Null", "'" & mstr˵�� & "'") & vbCrLf & _
                "Where ID = " & mlngClassID
        End If
    Case Val("1-������")
        If mlngGroupID = 0 Then
            mlngGroupID = GetNextID("zlRPTGroups")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlRPTGroups(ID,����ID,���,����,˵��) " & vbCrLf & _
                "Values(" & mlngGroupID & _
                IIF(lngClassID = 0, ",Null", "," & lngClassID) & _
                ",'" & mstr���� & "','" & mstr���� & "','" & mstr˵�� & "')"
        ElseIf Not (strOldName = mstr���� And strOld˵�� = mstr˵�� And lngClassID = mlngClassID) Then
            '˵�������Ʒ����仯
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Update zlRPTGroups " & vbCrLf & _
                "Set ���='" & mstr���� & "',����='" & mstr���� & "',˵��='" & mstr˵�� & "' " & vbCrLf & _
                "   ,����ID=" & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                "Where ID=" & mlngGroupID
            '����������̨�˵��ı������
            If mlngModule <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "Update zlPrograms " & vbCrLf & _
                    "Set ����='" & mstr���� & "',˵��='" & mstr˵�� & "' " & vbCrLf & _
                    "Where ���=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "Update zlMenus " & vbCrLf & _
                    "Set ����='" & mstr���� & "',�̱���='" & mstr���� & "',˵��='" & mstr˵�� & "' " & vbCrLf & _
                    "Where ID=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys
            End If
        End If
    Case Else
        'Ĭ��-����
        If mlngReportID = 0 Then
            '����
            If mlngSys <> 0 Then mlngSys = 0
            mlngReportID = GetNextID("zlReports")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlReports(ID,����ID,���,����,˵��,ϵͳ,�޸�ʱ��,����) " & vbCrLf & _
                "Values(" & mlngReportID & _
                "," & IIF(lngClassID = 0, "null", lngClassID) & _
                ",'" & mstr���� & "','" & mstr���� & "','" & mstr˵�� & "'," & vbCrLf & _
                IIF(mlngSys = 0, "NULL", mlngSys) & ",Sysdate," & AdjustStr(GetPass(mstr����, mstr����)) & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) " & vbCrLf & _
                "Values(" & mlngReportID & ",1,'" & mstr���� & "1'," & INIT_WIDTH & "," & INIT_HEIGHT & ",9,1,0,0)"

            '����������
            If fraGroups.Visible Then
                For i = 1 To vsfGroups.Rows - 1
                    If Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("ѡ��"))) <> 0 Then
                        lngTemp = Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("ID")))
                        strSQL = "Select Count(1) Rec From zlRPTSubs Where ��ID=[1]"
                        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, lngTemp)
                        If Not rsCheck.EOF Then
                            intOrder = Nvl(rsCheck!Rec, 0) + 1
                        Else
                            intOrder = 1
                        End If
                        rsCheck.Close
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlRPTSubs(��ID,����ID,���) " & vbCrLf & _
                            "Values(" & lngTemp & "," & mlngReportID & "," & intOrder & ")"
                        If mlngModule <> 0 Then
                            '����Ȩ�޼�¼
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) " & vbCrLf & _
                                "Values(" & IIF(mlngSys = 0, "NULL", mlngSys) & _
                                "," & mlngModule & _
                                ",'" & mstr���� & "'" & _
                                ",'" & mstr˵�� & "')"
                        End If
                    End If
                Next
            End If
        ElseIf Not (strOldCode = mstr���� And strOldName = mstr���� And strOld˵�� = mstr˵�� _
                        And lngClassID = mlngClassID And Val(vsfGroups.Tag) = 0) Then
            '�޸�
            If Not (strOldCode = mstr���� And strOldName = mstr���� And strOld˵�� = mstr˵�� _
                        And lngClassID = mlngClassID) Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "Update zlReports " & vbCrLf & _
                    "Set ���='" & mstr���� & "',����='" & mstr���� & "',˵��='" & mstr˵�� & "'" & vbCrLf & _
                    "   ,����=" & AdjustStr(GetPass(mstr����, mstr����)) & vbCrLf & _
                    "   ,����ID=" & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                    "Where ID=" & mlngReportID
                If mlngModule <> 0 Then '����������̨�˵��ı������
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = _
                        "Update zlPrograms " & vbCrLf & _
                        "Set ����='" & mstr���� & "',˵��='" & mstr˵�� & "' " & vbCrLf & _
                        "Where Upper(����)='ZL9REPORT' And ���=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = _
                        "Update zlMenus " & vbCrLf & _
                        "Set ����='" & mstr���� & "',�̱���='" & mstr���� & "',˵��='" & mstr˵�� & "' " & vbCrLf & _
                        "Where ģ��=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys & vbCrLf & _
                        "    And Exists(Select ���� From zlPrograms " & vbCrLf & _
                        "               Where Upper(����)='ZL9REPORT' And ���=" & mlngModule & " And Nvl(ϵͳ,0)=" & mlngSys & ")"
                End If
                
                '����������̨�ı������ӱ�Ĺ�����
                strSQL = _
                    "Select Distinct Nvl(B.ϵͳ, 0) ϵͳ, B.����id ���, a.��Id " & vbCrLf & _
                    "From Zlrptsubs a, Zlrptgroups b, Zlprograms c" & vbCrLf & _
                    "Where A.��id = B.Id And A.����id = [1]  And Nvl(B.ϵͳ, 0) = Nvl(C.ϵͳ, 0) " & vbCrLf & _
                    "    And B.����id = C.��� And Upper(C.����) = 'ZL9REPORT'"
                Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReportID)
                Do While Not rsCheck.EOF
                    If strOldName <> mstr���� Then  '�������Ʒ����仯
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '�����ӱ�������
                        arrSQL(UBound(arrSQL)) = _
                            "Update zlRPTSubs " & vbNewLine & _
                            "Set ���� = '" & mstr���� & "' " & vbNewLine & _
                            "Where ��Id = " & Nvl(rsCheck!��ID) & _
                            "    And ����Id = " & mlngReportID & " And ���� = '" & strOldName & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ϣ
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into Zlprogfuncs" & vbNewLine & _
                            "  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)" & vbNewLine & _
                            "  Select A.ϵͳ, A.���, '" & mstr���� & "', A.����, '" & mstr˵�� & "', A.ȱʡֵ" & vbNewLine & _
                            "  From Zlprogfuncs a" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "      And A.���� = '" & strOldName & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ȩ��Ϣ
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlrolegrant" & vbNewLine & _
                            "  (ϵͳ,���,��ɫ,����)" & vbNewLine & _
                            "  Select A.ϵͳ,A.���,A.��ɫ, '" & mstr���� & "' " & vbNewLine & _
                            "  From zlrolegrant a" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "     And A.���� = '" & strOldName & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ���ܶ���Ȩ����Ϣ
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlprogprivs" & vbNewLine & _
                            "  (ϵͳ,���,����,����,������,Ȩ��)" & vbNewLine & _
                            "  Select A.ϵͳ,A.���,'" & mstr���� & "',A.����,A.������,A.Ȩ��" & vbNewLine & _
                            "  From zlprogprivs a" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "      And A.���� = '" & strOldName & "'" & _
                            "      And Not Exists(Select 1 From zlProgPrivs " & vbCr & _
                            "                     Where Nvl(ϵͳ,0)=Nvl(a.ϵͳ,0) And ���=a.��� And ����='����' " & vbCr & _
                            "                         And ����=a.���� And ������=a.������ And Ȩ��=a.Ȩ��)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) 'ɾ��ԭʼ�����������ڴ��ڼ���ɾ����ϵ
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From Zlprogfuncs a " & vbNewLine & _
                            "Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "    And A.���� = '" & strOldName & "'"
                        'ϵͳ����š����� ����һ������Null������ɾ����ʧЧ
                        If Nvl(rsCheck!ϵͳ, 0) = 0 Or Nvl(rsCheck!���, 0) = 0 Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Delete From zlProgPrivs A " & vbNewLine & _
                                "Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                                "    And A.���� = '" & strOldName & "'"
                            
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Delete From zlRoleGrant A " & vbNewLine & _
                                "Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                                "    And A.���� = '" & strOldName & "'"
                        End If
                    Else '��������δ�����仯,ֻ����¹���˵��
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '���¹���˵��
                        arrSQL(UBound(arrSQL)) = _
                            "Update Zlprogfuncs A" & vbNewLine & _
                            "  Set A.˵��='" & mstr˵�� & "'" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "      And A.���� = '" & mstr���� & "'"
                    End If
                    rsCheck.MoveNext
                Loop
                
                '������ģ��ı�������
                strSQL = _
                    "Select Nvl(B.ϵͳ, 0) ϵͳ, B.����id ���, B.���� " & vbNewLine & _
                    "From Zlrptputs b, Zlprograms c, Zlprogfuncs d " & vbNewLine & _
                    "Where B.����id =[1] And Nvl(B.ϵͳ, 0) = Nvl(C.ϵͳ, 0) And B.����id = C.��� " & vbNewLine & _
                    "    And Upper(C.����) <> 'ZL9REPORT' And Nvl(C.ϵͳ, 0) = Nvl(D.ϵͳ, 0) And C.��� = D.��� " & vbNewLine & _
                    "    And D.���� = B.����"
                Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReportID)
                Do While Not rsCheck.EOF
                    If strOldName <> mstr���� And mlngSys = 0 Then   '��ϵͳ�������Ʒ����仯�����Զ����¹�������
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����Zlrptputs
                        arrSQL(UBound(arrSQL)) = _
                            "Update Zlrptputs Set ���� = '" & mstr���� & "' " & vbNewLine & _
                            "Where ����id = " & mlngReportID & " And Nvl(ϵͳ, 0) = " & rsCheck!ϵͳ & " And ����id = " & rsCheck!���
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ϣ
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into Zlprogfuncs" & vbNewLine & _
                            "  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)" & vbNewLine & _
                            "  Select A.ϵͳ, A.���, '" & mstr���� & "', A.����, '" & mstr˵�� & "', A.ȱʡֵ" & vbNewLine & _
                            "  From Zlprogfuncs a" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "     And A.���� = '" & rsCheck!���� & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ������Ȩ��Ϣ
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlrolegrant" & vbNewLine & _
                            "  (ϵͳ,���,��ɫ,����)" & vbNewLine & _
                            "  Select A.ϵͳ,A.���,A.��ɫ, '" & mstr���� & "' " & vbNewLine & _
                            "  From zlrolegrant a" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "      And A.���� = '" & rsCheck!���� & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '����һ��ԭʼ���ܶ���Ȩ����Ϣ
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlprogprivs" & vbNewLine & _
                            "  (ϵͳ,���,����,����,������,Ȩ��)" & vbNewLine & _
                            "  Select A.ϵͳ,A.���,'" & mstr���� & "',A.����,A.������,A.Ȩ��" & vbNewLine & _
                            "  From zlprogprivs a" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "      And A.���� = '" & rsCheck!���� & "'" & _
                            "      And Not Exists(Select 1 From zlProgPrivs " & vbCr & _
                            "                     Where ϵͳ=a.ϵͳ And ���=a.��� And ����='����' " & vbCr & _
                            "                         And ����=a.���� And ������=a.������ And Ȩ��=a.Ȩ��)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) 'ɾ��ԭʼ�����������ڴ��ڼ���ɾ����ϵ
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From Zlprogfuncs a " & vbNewLine & _
                            "Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "    And A.���� = '" & rsCheck!���� & "'"
                    Else '��ϵͳ����˵���仯���߹̶�����������ֻ���¹���˵��
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '���¹���˵��
                        arrSQL(UBound(arrSQL)) = _
                            "Update Zlprogfuncs A" & vbNewLine & _
                            "  Set A.˵��='" & mstr˵�� & "'" & vbNewLine & _
                            "  Where Nvl(A.ϵͳ, 0) = " & rsCheck!ϵͳ & " And A.��� = " & rsCheck!��� & vbNewLine & _
                            "     And A.���� = '" & rsCheck!���� & "'"
                    End If
                    rsCheck.MoveNext
                Loop
            End If
            
            '����������
            If fraGroups.Visible And Val(vsfGroups.Tag) = Val("1-�Ѳ���vsfGroup") Then
                For i = 1 To vsfGroups.Rows - 1
                    '��ȡ������ID
                    lngTemp = Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("ID")))
                    If Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("ѡ��"))) = 0 Then
                        '�Ƴ�������
                        If mlngModule <> 0 Then
                            '�ѷ������Զ��屨��
                            '����鱨�������Ȩ�޼�¼
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Delete zlProgFuncs " & _
                                "Where ϵͳ is null And ��� = " & mlngModule
                        Else
                            'δ�������Զ��屨��
                            lngProgID = ReportGroupIssue(lngTemp)
                            If lngProgID <> 0 Then
                                '�鱨���з���������ӱ����Ȩ�޼�¼
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = _
                                    "Delete zlProgFuncs " & _
                                    "Where ϵͳ is null And ��� = " & lngProgID & _
                                    "    And ���� = '" & mstr���� & "'"
                            End If
                        End If
                        
                        '�����ǰ������ӱ����¼
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Delete zlRPTSubs " & vbCrLf & _
                            "Where ��ID = " & lngTemp & " And ����ID =" & mlngReportID
                    Else
                        '���뱨����
                        
                        '��ȡ�ӱ�������
                        strSQL = "Select Count(1) Rec From zlRPTSubs Where ��ID=[1]"
                        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, lngTemp)
                        If Not rsCheck.EOF Then
                            intOrder = Nvl(rsCheck!Rec, 0) + 1
                        Else
                            intOrder = 1
                        End If
                        rsCheck.Close
                        
                        '�����ӱ����¼
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlRPTSubs(��ID,����ID,���,����) " & vbCrLf & _
                            "Select " & lngTemp & "," & mlngReportID & "," & intOrder & ",'" & mstr���� & "' " & vbCrLf & _
                            "From Dual " & vbCr & _
                            "Where Not Exists(Select 1 From zlRPTSubs " & vbCr & _
                            "                 Where ��ID = " & lngTemp & " And ����ID = " & mlngReportID & ")"
                        
                        If mlngModule <> 0 Then
                            '�ѷ������Զ��屨��
                            '����Ȩ�޼�¼
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) " & vbCrLf & _
                                "Select " & IIF(mlngSys = 0, "NULL", mlngSys) & _
                                "," & mlngModule & _
                                ",'" & mstr���� & "'" & _
                                ",'" & mstr˵�� & "'" & vbCrLf & _
                                "From Dual Where Not Exists(Select 1 From zlProgFuncs " & vbCrLf & _
                                "                           Where ϵͳ Is Null And ��� = " & mlngModule & ")"
                        Else
                            'δ�������Զ��屨��
                            lngProgID = ReportGroupIssue(lngTemp)
                            If lngProgID <> 0 Then
                                '�鱨���з������ӱ���ȱʡ�鱨��ĳ���ID�������ӱ����Ȩ�޼�¼��
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = _
                                    "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) " & vbCrLf & _
                                    "Select " & IIF(mlngSys = 0, "NULL", mlngSys) & _
                                    "," & lngProgID & _
                                    ",'" & mstr���� & "'" & _
                                    ",'" & mstr˵�� & "'" & vbCrLf & _
                                    "From Dual " & vbCr & _
                                    "Where Not Exists(Select 1 From zlProgFuncs " & vbCrLf & _
                                    "                 Where ϵͳ Is Null And ��� = " & lngProgID & _
                                    "                     And ���� = '" & mstr���� & "')"
                            End If
                        End If
                    End If
                Next
            End If
            
        End If
    End Select
    
    If UBound(arrSQL) >= 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            gcnOracle.Execute arrSQL(i)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    '�������
    Set grsReport = Nothing
    mblnOK = True
    Unload Me
    Exit Sub
    
hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume

    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mlngClassID <> 0 Or mlngGroupID <> 0 Or mlngReportID <> 0 Then txt����.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnViewGroups As Boolean
    
    mblnOK = False
    
    '��ȡ���������Ϣ
    If mbytMode = Val("2-������") Then
        Call mdlPublic.InitClass(cboClass, mlngClassID, mlngClassID)
    Else
        Call mdlPublic.InitClass(cboClass, mlngClassID)
    End If
    
    txt���.Text = mstr����
    txt����.Text = mstr����
    txt˵��.Text = mstr˵��
    
    Select Case mbytMode
    Case Val("2-������")
        If mlngClassID = 0 Then
            Caption = "�����������"
        Else
            Caption = "�޸ı������"
        End If
        
        '�ϼ�����
        For i = 0 To cboClass.ListCount - 1
            If mlngGroupID = Val(cboClass.ItemData(i)) Then
                cboClass.ListIndex = i
                Exit For
            End If
        Next
        
        cboClass.Enabled = mlngSys = 0
        
        lblClass.Left = 30
        lblClass.Caption = "�ϼ�����"
        lblName.Top = lblCode.Top
        lblDesc.Top = lblName.Top + 420
        txt����.Top = txt���.Top
        txt˵��.Top = txt���.Top + 420
        txt˵��.Height = txt˵��.Height + 420
        
        txt���.Text = ""
        txt���.Visible = False
        lblCode.Visible = False
    Case Val("1-������")
        If mlngGroupID = 0 Then
            Caption = "����������"
            txt���.Text = GetNextNO(True)
        Else
            Caption = "�޸ı�����"
        End If
        cboClass.Enabled = mlngSys = 0
    Case Val("3-�ӱ���")
        Caption = "�޸��ӱ���"
        cboClass.Enabled = False
        blnViewGroups = mlngSys = 0
    Case Else
        cboClass.Enabled = mlngSys = 0
        blnViewGroups = mlngSys = 0
        If mlngReportID = 0 Then
            Caption = "��������"
            txt���.Text = GetNextNO(False)
        Else
            Caption = "�޸ı���"
        End If
    End Select
    If mlngSys > 0 Then txt���.Enabled = False
    
    If blnViewGroups Then
        On Error GoTo hErr
        
        strSQL = _
            "Select Decode(Nvl(b.��id, 0), 0, 0, 1) ѡ��, a.���, a.����, a.Id " & vbNewLine & _
            "From zlRPTGroups A, zlRPTSubs B " & vbNewLine & _
            "Where a.Id = b.��id(+) And Nvl(a.ϵͳ, 0) = 0 And b.����Id(+) = [1] "
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ����ϵͳ�ı�������Ϣ", mlngReportID)
        
        With mvsfGroups
            .AppTemplate EM_Verify, vsfGroups, MSTR_GROUPS_COLS, "���|����"
            .Init True
            .Recordset = rsTemp
            .Repaint RT_Rows
        End With
        
        rsTemp.Close
    Else
        Me.Width = 4290
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    End If
    vsfGroups.Visible = blnViewGroups
    fraGroups.Visible = blnViewGroups
    
    Exit Sub
    
hErr:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mvsfGroups = Nothing
End Sub

Private Sub txt���_GotFocus()
    SelAll txt���
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf InStr(GSTR_SBC, Chr(KeyAscii)) > 0 Then
        KeyAscii = Asc(Mid(GSTR_DBC, InStr(GSTR_SBC, Chr(KeyAscii)), 1))
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If txt����.Text <> "" Then
        txt����.Text = ConvertSBC(txt����.Text)
    End If
End Sub

Private Sub txt˵��_GotFocus()
    SelAll txt˵��
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function GetClassID(ByVal lngID As Long, Optional ByVal blnGroup As Boolean = False) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    GetClassID = 0
    
    On Error GoTo hErr
    
    If blnGroup Then
        strSQL = "Select ����Id From zlRPTGroups Where Id = [1]"
    Else
        strSQL = "Select ����Id From zlReports Where Id = [1]"
    End If
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ�������ID", lngID)
    If rsTemp.EOF = False Then
        GetClassID = Nvl(rsTemp!����id, 0)
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Sub vsfGroups_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfGroups.Tag = "1"     '1��ʾ�Ѳ�������������ͨ���ñ�־���±���������
End Sub

Private Function ReportGroupIssue(ByVal lngID As Long) As Long
'���ܣ��жϱ������Ƿ��ѷ���
'������
'  lngID��������ID
'���أ�����0δ����������0�ѷ�������������ID��

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    ReportGroupIssue = 0
    
    strSQL = _
        "Select ����ID From zlRPTGroups Where ID = [1] And ����ʱ�� Is Not Null"
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "�жϱ������Ƿ��ѷ���", lngID)
    If rsTemp.RecordCount = 1 Then
        ReportGroupIssue = mdlPublic.Nvl(rsTemp!����id, 0)
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function
