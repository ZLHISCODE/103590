VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "frmSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6390
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6390
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   540
         TabIndex        =   6
         Top             =   60
         Width           =   90
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   165
         Picture         =   "frmSelect.frx":014A
         Top             =   30
         Width           =   240
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
      ScaleWidth      =   6390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3660
      Width           =   6390
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   765
         MaxLength       =   30
         TabIndex        =   10
         Top             =   120
         Width           =   1530
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4785
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3540
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   2640
         Top             =   240
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
      Begin VB.Label lblSearch 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   630
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
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2850
      Left            =   2520
      TabIndex        =   8
      Top             =   555
      Width           =   3765
      _cx             =   6641
      _cy             =   5027
      Appearance      =   0
      BorderStyle     =   0
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
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   30
      Cols            =   2
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
      ExplorerBar     =   1
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4905
      TabIndex        =   7
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
Public strSQLList As String
Public strSQLTree As String
Public strFLDList As String
Public strFLDTree As String
Public strParName As String '��������
Public bytType As Byte      '������������
Public strMatch As String '����ƥ�������
Public lngSeekHwnd As Long '���ڶ�λ����λ�õĿؼ�
Public mintConnect As Integer           '�������ӱ��

Public mblnMulti As Boolean '�Ƿ��ѡ��
Public mblnOK As Boolean
Public mlngSel As Long  '���е�ֵ�������ֵʱѡ��
Public mblnRelationReport As Boolean '�Ƿ��ѡ��������

'����δ����ʽ���������ԭʼֵ
Public strOutBand As String 'ѡ��İ�ֵ,��Ӧ&B
Public strOutDisp As String 'ѡ�����ʾֵ,��Ӧ&D

'�¸ĵĹ���������ر���
Public selectObjReport As Report
Public selectObjRelation As RPTRelations
Public selectlngType As Long
Public selectObjParent As Object
Public selectCurID As Integer

Private intPreNode As Long
Private blnItem As Boolean
Private blnSetFlex As Boolean, blnSetLvw As Boolean
Private rsList As ADODB.Recordset
Private strList As String
Private BlnSave As Boolean
Private rParent As RECT
Private mblnEnter As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim j As Long
    Dim blnFlag As Boolean
    Dim blnSelect As Boolean
    Dim strDisp As String, strBand As String
    
    On Error GoTo hErr
    
    strDisp = GetScript(strFLDList, "&D") '��ʾ���ֶ���
    strBand = GetScript(strFLDList, "&B") '�󶨵��ֶ���
    
    If strDisp = "" Or strBand = "" Then
        MsgBox "ѡ������û�ж��������İ󶨼���ʾ�ֶ���Ŀ��", vbInformation, App.Title
        Exit Sub
    End If
    
    If mblnRelationReport Or mblnMulti Then
        With vsf
            For i = 1 To vsf.Rows - 1
                blnFlag = False
                '��鹴ѡ�ı������������,���������1.��ѡ�˵�δ���ò�����2.�������˲�����ȡ����ѡ
                If Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 Then
                    If Split(.RowData(i), "|")(0) = "" Then
                        .Row = i
                        Call .ShowCell(i, .ColIndex("ѡ��"))
                        MsgBox "�������ݵ�""" & strDisp & """Ϊ��,����������""" & strParName & """����ʾ��", vbInformation, App.Title
                        Exit Sub
                    End If
                    If Split(.RowData(i), "|")(1) = "" Then
                        .Row = i
                        Call .ShowCell(i, .ColIndex("ѡ��"))
                        MsgBox "�������ݵ�""" & strBand & """Ϊ��,����������""" & strParName & """��󶨣�", vbInformation, App.Title
                        Exit Sub
                    End If
                    If mblnRelationReport Then
                        For j = 1 To selectObjRelation.count
                            If selectObjRelation.Item(j).��������ID = Split(.RowData(i), "|")(1) Then
                                blnFlag = True
                                Exit For
                            End If
                        Next
                        If blnFlag = False Then
                            '�ѹ�ѡ��δ���ò��������null
                            selectObjRelation.Add Val(Split(.RowData(i), "|")(1)), "NULL", "", CStr(Split(.RowData(i), "|")(0)), 0
                        End If
                    End If
                Else
                    If mblnRelationReport Then
                        'δ��ѡ���������˲���������ղ���
                        For j = 1 To selectObjRelation.count
                            If selectObjRelation.Item(j).��������ID = Split(.RowData(i), "|")(1) Then
                                selectObjRelation.Remove j
                            End If
                        Next
                    End If
                End If
            Next
            '������ʾ��,�󶨴�
            strOutDisp = ""
            strOutBand = ""
            For i = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 Or .Rows = 2 Then
                    strOutDisp = strOutDisp & "," & Split(.RowData(i), "|")(0)
                    strOutBand = strOutBand & "," & Split(.RowData(i), "|")(1)
                End If
                If mblnRelationReport Then
                    If Abs(Val(.TextMatrix(i, .ColIndex("Ĭ��")))) = 1 And Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 Then
                        For j = 1 To selectObjRelation.count
                            If selectObjRelation.Item(j).��������ID = Split(.RowData(i), "|")(1) Then
                                selectObjRelation.Item(j).Ĭ�� = 1
                            End If
                        Next
                        blnSelect = True
                    ElseIf Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 And Abs(Val(.TextMatrix(i, .ColIndex("Ĭ��")))) = 0 Then
                        For j = 1 To selectObjRelation.count
                            If selectObjRelation.Item(j).��������ID = Split(.RowData(i), "|")(1) Then
                                selectObjRelation.Item(j).Ĭ�� = 0
                            End If
                        Next
                    End If
                End If
            Next
            If Trim(strOutDisp) = "" Or Trim(strOutBand) = "" Then
                If mblnRelationReport Then
                    For j = 1 To selectObjRelation.count
                        selectObjRelation.Remove j
                    Next
                    GoTo endHand
                End If
                If mblnMulti Then
                    MsgBox "û��ѡ���κ����ݣ�", vbInformation, App.Title
                    If vsf.Enabled And vsf.Visible Then vsf.SetFocus
                    Exit Sub
                End If
                
            ElseIf UBound(Split(strOutBand, ",")) > 1000 Then
                MsgBox "ѡ������ݹ��࣡", vbInformation, App.Title
                Exit Sub
            End If
            If mblnRelationReport Then
                If blnSelect = False Then
                    MsgBox "û��ѡ��Ĭ�Ϲ����ı����빴ѡ��", vbInformation, App.Title
                    Exit Sub
                End If
            End If
            strOutDisp = Mid(strOutDisp, 2)
            If LCase(TypeName(selectObjParent)) = "frmdata" Then
                strOutBand = Mid(strOutBand, 2)
            Else
                strOutBand = " IN (" & Mid(strOutBand, 2) & ") "
            End If
        End With
    Else
        If vsf.Row = -1 Then
            MsgBox "û��ѡ���κ����ݣ�", vbInformation, App.Title
            Exit Sub
        End If
        If InStr(vsf.RowData(vsf.Row), "|") <= 0 Then
            MsgBox "�������ݵ�Ϊ�գ���������Դ��", vbInformation, App.Title
            Exit Sub
        End If
        If Split(vsf.RowData(vsf.Row), "|")(0) = "" Then
            MsgBox "�������ݵ�""" & strDisp & """Ϊ��,����������""" & strParName & """����ʾ��", vbInformation, App.Title
            Exit Sub
        End If
        If Split(vsf.RowData(vsf.Row), "|")(1) = "" Then
            MsgBox "�������ݵ�""" & strBand & """Ϊ��,����������""" & strParName & """��󶨣�", vbInformation, App.Title
            Exit Sub
        End If
        
        '���ͼ��
        Select Case bytType
            Case 1
                If Not IsNumeric(Split(vsf.RowData(vsf.Row), "|")(1)) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    If vsf.Enabled And vsf.Visible Then vsf.SetFocus
                    Exit Sub
                End If
            Case 2
                If Not IsDate(Split(vsf.RowData(vsf.Row), "|")(1)) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    If vsf.Enabled And vsf.Visible Then vsf.SetFocus
                    Exit Sub
                End If
        End Select
    
        strOutDisp = Split(vsf.RowData(vsf.Row), "|")(0)
        strOutBand = Split(vsf.RowData(vsf.Row), "|")(1)
        
    End If
    
endHand:
    mblnOK = True
    On Error Resume Next
    Hide
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub Form_Activate()
    If Me.Visible Then
        If vsf.Rows > 1 And mblnEnter Then
            vsf.Row = 1: vsf.Col = 0
            If vsf.Visible And vsf.Enabled Then vsf.SetFocus
            mblnEnter = False
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Not vsf.Visible Then Exit Sub
        
        For i = 1 To vsf.Rows - 1
            If vsf.ColIndex("ѡ��") >= 0 Then
                vsf.TextMatrix(i, vsf.ColIndex("ѡ��")) = 1
            End If
        Next
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Not vsf.Visible Then Exit Sub
        
        For i = 1 To vsf.Rows - 1
            If vsf.ColIndex("ѡ��") >= 0 Then
                vsf.TextMatrix(i, vsf.ColIndex("ѡ��")) = 0
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim lngW As Long, i As Long
    Dim strDataSourceName As String
    Dim X As Long, Y As Long, k As Long
    
    mblnEnter = False
    If Not InDesign Then
        glngSelProc = GetWindowLong(hwnd, GWL_WNDPROC)
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SelMessage)
    End If
    
    mblnOK = False
    BlnSave = True
    blnSetFlex = False '�Ƿ��Ѿ��Ա��ָ����
    blnSetLvw = False
    intPreNode = 0
    
    strOutBand = ""
    strOutDisp = ""
    
    Me.Caption = strParName & "ѡ����"
    
    strSQLList = Replace(strSQLList, "[*]", strMatch)
    strSQLTree = Replace(strSQLTree, "[*]", strMatch)
    If mblnRelationReport = False Then
        lblSearch.Visible = False
        txtSearch.Visible = False
    End If
    If strSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
        If Not FillList Then BlnSave = False: Unload Me: Exit Sub
    Else
        tvw_s.Visible = True
        If Not FillTree Then BlnSave = False: Unload Me: Exit Sub
        If tvw_s.Nodes.count > 0 Then
            tvw_s.Nodes(1).Selected = True
            If Not tvw_s.Nodes(1).Child Is Nothing And strMatch = "" Then
                tvw_s.Nodes(1).Child.Selected = True
            End If
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        End If
    End If

    '����ƥ���Զ�����
    If strMatch <> "" Then
        If rsList.RecordCount = 1 Then
            If mblnRelationReport Then
                '�����ù�������
                With vsf
                    .Row = 1
                    If selectlngType = 2 Then
                        X = InStr(1, selectObjReport.Items("_" & selectCurID).����, "]")
                        Y = InStr(1, selectObjReport.Items("_" & selectCurID).����, ".")
                        k = InStr(1, selectObjReport.Items("_" & selectCurID).����, "[")
                        If X > k And X > Y And X <> 0 And k <> 0 Then
                            Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), _
                                Mid(selectObjReport.Items("_" & selectCurID).����, k + 1, Y - k - 1), selectObjReport, _
                                Split(.RowData(.Row), "|")(0), selectObjRelation, 2)
                        End If
                    ElseIf selectlngType = 3 Then
                        Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), _
                            frmData.txtRelation.Text, selectObjReport, Split(.RowData(.Row), "|")(0), _
                            selectObjRelation, 3)
                    Else
                        If LCase(TypeName(selectObjParent)) = "frmformula" Then
                            X = InStr(1, frmFormula.txtFormula.Text, "]")
                            Y = InStr(1, frmFormula.txtFormula.Text, ".")
                            k = InStr(1, frmFormula.txtFormula.Text, "[")
                            If X > k And X > Y And X <> 0 And k <> 0 Then
                                Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), Mid(frmFormula.txtFormula.Text, k + 1, Y - k - 1), selectObjReport, Split(.RowData(.Row), "|")(0), selectObjRelation, 1)
                            Else
                                MsgBox "��ǰ�б����Ȱ�һ������Դ�����磺[����Դ.�ֶ�],�󶨺������ù�������", vbInformation, Me.Caption
                            End If
                        ElseIf LCase(TypeName(selectObjParent)) = "frmdata" Then
                            '��ȡ����Դ������
                            If selectObjParent.tvw.Nodes.count > 0 Then
                                strDataSourceName = selectObjParent.tvw.Nodes(1).Text
                                Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), _
                                    strDataSourceName, selectObjReport, Split(.RowData(.Row), "|")(0), _
                                    selectObjRelation, 1)
                            End If
                        End If
                    End If
                End With
            
                If vsf.ColIndex("ѡ��") >= 0 Then
                    vsf.TextMatrix(1, vsf.ColIndex("ѡ��")) = 1
                End If
                If mblnRelationReport Then
                    vsf.TextMatrix(1, vsf.ColIndex("Ĭ��")) = 1
                End If
            End If
            Call cmdOK_Click
            BlnSave = False
            Unload Me: Exit Sub
        ElseIf rsList.RecordCount = 0 Then
            MsgBox "û���ҵ���ƥ�����Ŀ,���������룡", vbInformation, App.Title
            BlnSave = False
            Call cmdCancel_Click: Exit Sub
        End If
    End If
    
    Call Form_Resize
    
    '���弰�б�ȱʡ���
    For i = 0 To vsf.Cols - 1
        lngW = lngW + vsf.ColWidth(i)
    Next
    Me.Width = lngW + 500 + IIF(strSQLTree = "", 0, tvw_s.Width + pic.Width)
    If Me.Width < 3000 Then Me.Width = 3000
    If strSQLTree <> "" Then
        If Me.Width < (tvw_s.Width + pic.Width) * 2.2 Then Me.Width = (tvw_s.Width + pic.Width) * 2.2
    End If
    
    If Me.Width < 7000 Then Me.Width = 7000
    
    RestoreWinState Me, App.ProductName, strParName
    
    If strSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
    Else
        tvw_s.Visible = True
    End If
    
    If vsf.ColIndex("ѡ��") >= 0 Then
        vsf.ToolTipText = "ȫѡ(Ctrl+A),ȫ��(Ctrl+R)"
    Else
        vsf.ToolTipText = ""
    End If
    
    '��λ
    If lngSeekHwnd <> 0 Then
        Call Form_Resize
        GetWindowRect lngSeekHwnd, rParent
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
    mblnEnter = True
End Sub

Private Sub Form_Resize()
    Dim lngTVW As Long
    
    On Error Resume Next
    
    lngTVW = IIF(tvw_s.Visible, tvw_s.Width + pic.Width, 0)
    
    tvw_s.Left = Me.ScaleLeft
    tvw_s.Top = picInfo.Top + picInfo.Height + 15
    tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height - 15
    
    pic.Left = tvw_s.Left + tvw_s.Width
    pic.Top = tvw_s.Top
    pic.Height = tvw_s.Height
    
    vsf.Left = Me.ScaleLeft + lngTVW
    vsf.Top = tvw_s.Top
    vsf.Height = tvw_s.Height
    vsf.Width = Me.ScaleWidth - lngTVW
    
    lbl.Left = vsf.Left
    lbl.Top = vsf.Top
    lbl.Width = vsf.Width
    lbl.Height = vsf.Height
    
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 45
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strMatch = ""
    lngSeekHwnd = 0
    If BlnSave Then SaveWinState Me, App.ProductName, strParName
    If Not InDesign Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngSelProc)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsf.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        
        vsf.Left = vsf.Left + X
        vsf.Width = vsf.Width - X
        
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
    Dim i As Long
    Dim objNode As Node
    Dim strSel As String, strRela As String
    
    On Error GoTo errH
    
    strSel = GetScript(strFLDTree, "&S")
    strRela = GetScript(strFLDTree, "&R")
    
    If strSel = "" Or strRela = "" Then
        MsgBox "δ��������ѡ�������ϸ�б���������ֶ���Ŀ��", vbInformation, App.Title
        Exit Function
    End If
    Call OpenRecord(rsTmp, RemoveNote(strSQLTree), Me.Caption & "_FillTree", mintConnect) 'SQLһ��̶�,[*]��SQL��''��,�����޷�����
    
    tvw_s.Nodes.Clear
        
    If InStr("|" & UCase(strFLDTree), "|ID,") > 0 And InStr("|" & UCase(strFLDTree), "|�ϼ�ID,") > 0 Then
        '���������б���ʾ
        Set objNode = tvw_s.Nodes.Add(, , "ALL", "������Ŀ", 1)
        objNode.Tag = "ALL"
        objNode.Expanded = True
        
        For i = 1 To rsTmp.RecordCount
            If IsNull(rsTmp!�ϼ�ID) Then
                Set objNode = tvw_s.Nodes.Add("ALL", 4, "_" & rsTmp!id, IIF(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
            Else
                Set objNode = tvw_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, IIF(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
            End If
            objNode.Tag = IIF(IsNull(rsTmp.Fields(strRela).Value), "", rsTmp.Fields(strRela).Value)
            rsTmp.MoveNext
        Next
    Else
        '����һ���б���ʾ
        For i = 1 To rsTmp.RecordCount
            Set objNode = tvw_s.Nodes.Add(, , , IIF(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
            objNode.Tag = IIF(IsNull(rsTmp.Fields(strRela).Value), "", rsTmp.Fields(strRela).Value)
            rsTmp.MoveNext
        Next
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
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=" & strKey
                    Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "='" & strKey & "'"
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
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

Private Function HaveScript(strFld As String, strName As String, strType As String) As Boolean
'���ܣ��ж����ֶ������У�ָ�����ֶ��Ƿ����ָ������������
'������strName=�ֶ���,strFld=�ֶ�������,strType="&S &D &B &R"
'���أ�False=δ�����ֶλ��ֶβ�����ָ������
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If Split(Split(strFld, "|")(i), ",")(0) = strName Then
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
    Dim strSQL As String, strValue As String
    Dim i As Long, j As Long
    Dim objItem As ListItem
    Dim strDisp As String, strBand As String
    Dim blnSelect As Boolean
    
    On Error GoTo errH
    
    vsf.Rows = 1
    vsf.Cols = 0
    
    '����Ϊֻ��������
    If Not blnSort Then
        If strSQLTree = "" Then
            strSQL = strSQLList
        Else
            '��̬����ϸ���ݴ���Ϊֻ��ȡ�����ķ��ಿ��(���� Order by �Ӿ�)
            If strKey = "ALL" Then
                strSQL = strSQLList
            Else
                strSQL = GetRelaSQL(strSQLList, strFLDList, strKey)
            End If
            
            If strSQL = "" Then
                MsgBox "�������ݶ�ȡʧ�ܣ�", vbInformation, App.Title
                Exit Function
            End If
        End If
        
        Screen.MousePointer = 11
        Me.Refresh
        
        Set rsList = New ADODB.Recordset
        Call OpenRecord(rsList, RemoveNote(strSQL), Me.Caption & "_FillList", mintConnect)     'SQLһ��̶�,[*]��SQL��''��,�����޷�����
    End If
    
    If Not rsList.EOF Then
        AddListCols
        
        strDisp = GetScript(strFLDList, "&D") '��ʾֵ��Ŀ
        strBand = GetScript(strFLDList, "&B") '��ֵ��Ŀ
        
        For i = 1 To rsList.RecordCount
            
            With vsf
                If mblnRelationReport Or mblnMulti Then
                    '��һ��ѡ�����Դӵ�2�п�ʼ
                    strValue = GetValue(rsList.Fields(vsf.ColKey(1)))
                    If .ColData(0) <> "" Then strValue = Format(strValue, .ColData(1))
                Else
                    strValue = GetValue(rsList.Fields(vsf.ColKey(0)))
                    If .ColData(0) <> "" Then strValue = Format(strValue, .ColData(0))
                End If
                
                .Rows = .Rows + 1
                For j = IIF(mblnRelationReport Or mblnMulti, 1, 0) To .Cols - 1
                    If Not (mblnRelationReport And (.ColKey(j) = "ѡ��" Or .ColKey(j) = "Ĭ��" Or .ColKey(j) = "����")) Then
                        strValue = GetValue(rsList.Fields(.ColKey(j)))
                        If .ColData(j) <> "" Then strValue = Format(strValue, .ColData(j))
                        .TextMatrix(.Rows - 1, j) = strValue
                    End If
                Next
                
                '����ʾֵ����ֵ������TAG��,��Ϊ��һ����Щ�ֶλ�Ϊѡ���ֶ�
                '��ʽΪ"��ʾֵ|��ֵ"
                If strDisp <> "" Then
                    .RowData(.Rows - 1) = IIF(IsNull(rsList.Fields(strDisp).Value), "", rsList.Fields(strDisp).Value)
                End If
                .RowData(.Rows - 1) = .RowData(.Rows - 1) & "|"
                If strBand <> "" Then
                    .RowData(.Rows - 1) = .RowData(.Rows - 1) & IIF(IsNull(rsList.Fields(strBand).Value), "", rsList.Fields(strBand).Value)
                    If mblnRelationReport Then
                        If CollectionHave(Val(rsList.Fields(strBand).Value & "")) Then
                            If vsf.ColIndex("ѡ��") >= 0 Then
                                .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 1
                            End If
                            If blnSelect = False Then
                                .Row = .Rows - 1
                                Call .ShowCell(.Row, 1)
                            End If
                            blnSelect = True
                        End If
                        If mblnRelationReport Then
                            If CheckSelect(Val(rsList.Fields(strBand).Value & "")) Then
                                .TextMatrix(.Rows - 1, .ColIndex("Ĭ��")) = 1
                            End If
                        End If
                    End If
                    If mlngSel <> 0 And Val(rsList.Fields(strBand).Value & "") = mlngSel Then .Row = .Rows - 1
                End If
                
                '����������Ϊ...
                If mblnRelationReport Then
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = "����ֵ"
                End If
            End With
            '------------------------------------------------------------------------------------------
            rsList.MoveNext
        Next
        If vsf.Rows > 1 And vsf.Row <= 0 Then vsf.Row = 1
        '�Զ������п�
        vsf.AutoResize = True
        Call vsf.AutoSize(0, vsf.Cols - 1)
        '��Ϊ��ѡ�����µ����п�
        If mblnRelationReport Then
            vsf.ColWidth(vsf.ColIndex("ѡ��")) = 600
            vsf.ColWidth(vsf.ColIndex("Ĭ��")) = 600
            vsf.ColWidth(vsf.ColIndex("����")) = 900
        End If
        If mblnRelationReport Then
            lblInfo.Caption = "����" & rsList.RecordCount & "�ű���"
        Else
            lblInfo.Caption = "����" & rsList.RecordCount & "�����"
        End If
    Else
        'û������ʱ����ʾ�յ�vsf(����ͷ)
        Call AddListCols
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

Private Function CollectionHave(ByVal lngKey As String) As Boolean
    On Error GoTo ErrHand
    '���ܣ����ҵ�ǰID�Ƿ�����ڹ���������
    Dim i As Integer
    For i = 1 To selectObjRelation.count
        If Val(selectObjRelation.Item(i).��������ID) = Val(lngKey) Then
            CollectionHave = True
            Exit For
        End If
    Next
    Exit Function
ErrHand:
    '�����ڷ���False
    If Err.Number = 5 Then CollectionHave = False
    Err.Clear
End Function

Private Function CheckSelect(ByVal lngKey As Long) As Boolean
    On Error GoTo ErrHand
    '���ܣ����ҵ�ǰID�Ƿ�����ڹ���������
    Dim i As Integer
    For i = 1 To selectObjRelation.count
        If Val(selectObjRelation.Item(i).��������ID) = lngKey Then
            If Val(selectObjRelation.Item(i).Ĭ��) = 1 Then
                CheckSelect = True
                Exit For
            End If
        End If
    Next
    Exit Function
ErrHand:
    '�����ڷ���False
    If Err.Number = 5 Then CheckSelect = False
    Err.Clear
End Function

Private Sub AddListCols()
'���ܣ�����strFLDList�ֶ�����ֵ,ΪListView������ͷ
    Dim i As Integer, j As Integer
    Dim strFld As String
    Dim objCol As ColumnHeader
    Dim intCol As Integer

    With vsf
        If mblnRelationReport Or mblnMulti = True Then
            intCol = 0
            .Cols = 1
            .TextMatrix(0, 0) = "ѡ��"
            .ColKey(0) = "ѡ��"
            .ColDataType(0) = flexDTBoolean
            .ColWidth(0) = 600
            intCol = intCol + 1
            .Editable = flexEDKbdMouse
        End If
        For i = 0 To UBound(Split(strFLDList, "|"))
            strFld = Split(strFLDList, "|")(i)
            If strFld Like "*&S*" Then
                
                .Cols = intCol + 1
                .ColKey(intCol) = Split(strFld, ",")(0)
                .TextMatrix(0, intCol) = Split(strFld, ",")(0)
                .ColWidth(intCol) = Me.TextWidth(Split(strFld, ",")(0) & "��")
                
                '�����ֶ������������ö���(��1ֻ�������)
                Select Case Split(strFld, ",")(1)
                    Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                        If rsList.Fields(.TextMatrix(0, intCol)).NumericScale > 0 Then
                            j = rsList.Fields(.TextMatrix(0, intCol)).NumericScale
                            .ColData(intCol) = "0." & String(IIF(j > 2, 2, j), "0; ;")
                            If intCol <> 1 Then .ColAlignment(intCol) = flexAlignRightCenter
                        ElseIf intCol <> 1 Then
                            If rsList.Fields(.TextMatrix(0, intCol)).Precision < 3 Then
                                .ColAlignment(intCol) = flexAlignCenterCenter
                            Else
                                .ColAlignment(intCol) = flexAlignLeftCenter
                            End If
                        End If
                        If .TextMatrix(0, intCol) Like "*��" Then .ColData(intCol) = "0.000"
                        If .TextMatrix(0, intCol) Like "*��" Then .ColData(intCol) = "0.00"
                    Case adDBTimeStamp, adDBTime, adDBDate, adDate
                        If intCol <> 1 Then .ColAlignment(intCol) = flexAlignLeftCenter
                    Case Else
                        If intCol <> 1 Then .ColAlignment(intCol) = flexAlignLeftCenter
                End Select
                If .TextMatrix(0, intCol) Like "*��λ*" And intCol <> 1 Then .ColAlignment(intCol) = flexAlignCenterCenter
                If .TextMatrix(0, intCol) Like "*��*" And intCol <> 1 Then .ColAlignment(intCol) = flexAlignCenterCenter
                intCol = intCol + 1
            End If
        Next
        If mblnRelationReport Then
            .Cols = .Cols + 2
            .TextMatrix(0, intCol) = "Ĭ��"
            .TextMatrix(0, intCol + 1) = "����"
            .ColKey(intCol) = "Ĭ��"
            .ColKey(intCol + 1) = "����"
            .ColDataType(intCol) = flexDTBoolean
            .ColWidth(intCol) = 600
            .ColWidth(intCol + 1) = 600
            .ColComboList(.ColIndex("����")) = "..."
        End If
    End With
End Sub

Private Function GetValue(objFld As Field) As String
'����:�����ֶ�����ȡ���ʵ���ʾֵ
    Dim strValue As String
    Select Case objFld.type
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            strValue = IIF(IsNull(objFld.Value), 0, objFld.Value)
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
            If Format(strValue, "HH:mm:ss") = "00:00:00" Then
                strValue = Format(strValue, "yyyy-MM-dd")
            Else
                strValue = Format(strValue, "yyyy-MM-dd HH:mm:ss")
            End If
        Case Else
            strValue = IIF(IsNull(objFld.Value), "", objFld.Value)
    End Select
    GetValue = strValue
End Function

Private Sub txtSearch_GotFocus()
    txtSearch.SelStart = 0: txtSearch.SelLength = Len(txtSearch.Text)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "%" Or Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "*" Then KeyAscii = 0
    If Trim(txtSearch.Text) = "" Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        Dim strKey As String
        
        strKey = txtSearch.Text
        FindLocal strKey
    End If
    txtSearch.SetFocus
End Sub

Private Sub FindLocal(ByVal strKey As String, Optional blnResume As Boolean)
    Dim lngStart As Long
    Dim i As Long
    Dim blnFind As Boolean
    
    With vsf
        If blnResume Then lngStart = 1 Else lngStart = .Row + 1
        If lngStart > .Rows - 1 Then lngStart = 1
        For i = IIF(lngStart = 0, 1, lngStart) To .Rows - 1
            If Split(.RowData(i), "|")(0) Like "*" & strKey & "*" Or .TextMatrix(i, .ColIndex("���")) Like "*" & strKey & "*" Then
                Call .ShowCell(i, 1)
                .Row = i
                blnFind = True
                Exit For
            End If
        Next
        
        If blnFind = False Then
            If MsgBox(" �Ѿ���λ�������ҵ�����Ϣ���Ƿ����²��ң�", vbInformation + vbYesNo, App.Title) = vbYes Then
                Call FindLocal(strKey, True)
            Else
                txtSearch.SetFocus
                Call txtSearch_GotFocus
            End If
        End If
    End With
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case vsf.ColIndex("ѡ��")
        If Abs(Val(vsf.TextMatrix(Row, Col))) = 0 And vsf.ColIndex("Ĭ��") >= 0 Then
            vsf.TextMatrix(Row, vsf.ColIndex("Ĭ��")) = 0
        End If
    Case vsf.ColIndex("Ĭ��")
        If vsf.ColIndex("ѡ��") >= 0 Then
            If Abs(Val(vsf.TextMatrix(Row, vsf.ColIndex("ѡ��")))) = 0 Then
                vsf.TextMatrix(Row, Col) = 0
            End If
        End If
    End Select
End Sub

Private Sub vsf_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnRelationReport Then
        If Col <> vsf.ColIndex("����") And Col <> vsf.ColIndex("ѡ��") And Col <> vsf.ColIndex("Ĭ��") Then Cancel = True
        If vsf.Col = vsf.ColIndex("Ĭ��") Then
            If Abs(Val(vsf.TextMatrix(Row, vsf.ColIndex("ѡ��")))) = 0 Then Cancel = True
        End If
    End If
    If mblnMulti Then
        If Col <> vsf.ColIndex("ѡ��") Then Cancel = True
    End If
    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim X As Long, Y As Long, k As Long
    Dim strDataSourceName As String
    
    With vsf
        Select Case Col
        Case .ColIndex("����")
            If selectlngType = 2 Then
                X = InStr(1, selectObjReport.Items("_" & selectCurID).����, "]")
                Y = InStr(1, selectObjReport.Items("_" & selectCurID).����, ".")
                k = InStr(1, selectObjReport.Items("_" & selectCurID).����, "[")
                If X > k And X > Y And X <> 0 And k <> 0 Then
                    Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), _
                        Mid(selectObjReport.Items("_" & selectCurID).����, k + 1, Y - k - 1), _
                        selectObjReport, Split(.RowData(.Row), "|")(0), selectObjRelation, 2)
                End If
            ElseIf selectlngType = 3 Then
                Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), _
                    frmData.txtRelation.Text, selectObjReport, Split(.RowData(.Row), "|")(0), _
                    selectObjRelation, 3)
            Else
                If LCase(TypeName(selectObjParent)) = "frmformula" Then
                    X = InStr(1, selectObjParent.txtFormula.Text, "]")
                    Y = InStr(1, selectObjParent.txtFormula.Text, ".")
                    k = InStr(1, selectObjParent.txtFormula.Text, "[")
                    If X > k And X > Y And X <> 0 And k <> 0 Then
                        Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), _
                            Mid(frmFormula.txtFormula.Text, k + 1, Y - k - 1), selectObjReport, _
                            Split(.RowData(.Row), "|")(0), selectObjRelation, 1)
                    Else
                        MsgBox "��ǰ�б����Ȱ�һ������Դ�����磺[����Դ.�ֶ�],�󶨺������ù�������", vbInformation, Me.Caption
                    End If
                ElseIf LCase(TypeName(selectObjParent)) = "frmdata" Then
                    '��ȡ����Դ������
                    If selectObjParent.tvw.Nodes.count > 0 Then
                        strDataSourceName = selectObjParent.tvw.Nodes(1).Text
                        Call frmRelationSetup.ShowMe(selectObjParent, Val(Split(.RowData(.Row), "|")(1)), _
                            strDataSourceName, selectObjReport, Split(.RowData(.Row), "|")(0), _
                            selectObjRelation, 1)
                    End If
                End If
            End If
        End Select
    End With
End Sub

Private Sub vsf_Click()
    Dim i As Long
    
    If vsf.Row < 1 Then Exit Sub
    If vsf.Col = vsf.ColIndex("Ĭ��") Then
        If vsf.Col = vsf.ColIndex("Ĭ��") Then
            If Abs(Val(vsf.TextMatrix(vsf.Row, vsf.ColIndex("ѡ��")))) = 0 Then Exit Sub
        End If
        For i = 1 To vsf.Rows - 1
            vsf.TextMatrix(i, vsf.ColIndex("Ĭ��")) = 0
        Next
        vsf.TextMatrix(vsf.Row, vsf.Col) = 1
    End If
End Sub

Private Sub vsf_DblClick()
    Call vsf_CellButtonClick(vsf.Row, vsf.Col)
End Sub
