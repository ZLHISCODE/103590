VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmDocShiftBase 
   Caption         =   "ҽ�����Ӱ���������"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13620
   Icon            =   "frmDocShiftBase.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13620
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsfPatiTypeInfo 
      Height          =   6975
      Left            =   3360
      TabIndex        =   1
      Top             =   1200
      Width           =   9255
      _cx             =   16325
      _cy             =   12303
      Appearance      =   0
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
      BackColorBkg    =   -2147483643
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDocShiftBase.frx":5C02
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPatiType 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3075
      _cx             =   5424
      _cy             =   12303
      Appearance      =   0
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
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   400
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDocShiftBase.frx":5EA5
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8295
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDocShiftBase.frx":5F96
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21114
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":682A
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":6DC4
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":735E
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":DBC0
            Key             =   "add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":14422
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":14E34
            Key             =   "CUp"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":152CE
            Key             =   "CMid"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":15768
            Key             =   "CDown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":15C02
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":16614
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":17026
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   2640
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmDocShiftBase.frx":1D888
      Left            =   3600
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDocShiftBase.frx":1D89C
   End
   Begin VB.Label lblPatiTypeInfo 
      AutoSize        =   -1  'True
      Caption         =   "������Ŀ"
      Height          =   180
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblPatiType 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "frmDocShiftBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mobjView As Object

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strSName As String, strPatiPrj As String
    Dim objControl As CommandBarControl
    Dim i As Long
    
    With vsfPatiType
        If .Rows > 1 Then
            strSName = .TextMatrix(.Row, .ColIndex("���"))
        End If
    End With
    With vsfPatiTypeInfo
        If .Rows > 1 Then
            strPatiPrj = .TextMatrix(.Row, .ColIndex("��Ŀ����"))
        End If
    End With
    Select Case Control.ID
        Case conMenu_DocShift_File_Preview
            Call PreView
        Case conMenu_DocShift_Edit_New
            Call PatiType(1, strSName)
        Case conMenu_DocShift_Edit_Modify
            Call PatiType(2, strSName)
        Case conMenu_DocShift_Edit_Delete
            If MsgBox("��ȷ��ɾ�����Ϊ��" & strSName & "���Ĳ���������", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
            On Error GoTo errH
            gstrSql = "Zl_ҽ�����Ӱಡ������_Edit(3,'" & strSName & "')"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            With vsfPatiType
                .RemoveItem .Row
                If .Row > 0 Then Call vsfPatiType_Click
            End With
        Case conMenu_DocShift_Edit_Reuse
            Call ReuseStop(4, strSName)
        Case conMenu_DocShift_Edit_Stop
            Call ReuseStop(5, strSName)
        Case conMenu_DocShift_Edit_NewProject
            Call PatiPrj(1, strSName, strPatiPrj)
        Case conMenu_DocShift_Edit_ModifyProject
            Call PatiPrj(2, strSName, strPatiPrj)
        Case conMenu_DocShift_Edit_DeleteProject
            If MsgBox("��ȷ��ɾ����Ŀ����Ϊ��" & strPatiPrj & "���Ĳ���������", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
            On Error GoTo errH
            gstrSql = "Zl_ҽ�����Ӱಡ����Ŀ_Edit(3,'" & strSName & "','" & strPatiPrj & "')"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            vsfPatiTypeInfo.RemoveItem vsfPatiTypeInfo.Row
        Case conMenu_DocShift_Edit_RowProject
            '(ȡ��)�ϲ���ʵ�ʾ��Ǹı����
            Call AdjustNum
        Case conMenu_View_ToolBar_Button '������
            For i = 2 To cbsMain.Count
                Control.Checked = Not Control.Checked
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Call Form_Resize
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '��ť����
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Call Form_Resize
            Me.cbsMain.RecalcLayout
        Case conMenu_DocShift_Help_Web_Home
            Call zlHomePage(Me.hwnd)
        Case conMenu_DocShift_Help_Web_Mail
            Call zlMailTo(Me.hwnd)
        Case conMenu_DocShift_Help_About
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_DocShift_File_Exit
            Unload Me
    End Select
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub PreView()
'Ԥ������Ч��
    If CreateObj(mobjView) Then
        Call mobjView.ShowViewShift(Me, vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("���")))
    End If
End Sub

Private Function CreateObj(ByRef objView As Object) As Boolean
'��������
        
    If objView Is Nothing Then
        On Error Resume Next
        Set objView = CreateObject("zl9DoctorShift.clsDoctorShift")
        err.Clear: On Error GoTo 0
        If objView Is Nothing Then
            MsgBox "zl9DoctorShift����δ�����ɹ���", vbInformation, gstrSysName
            Exit Function
        Else
            Call objView.InitDoctorShift(glngSys, gcnOracle)
        End If
    End If
    CreateObj = True
End Function

Private Sub ReuseStop(ByVal bytType As Byte, ByVal strName As String)
'bytType:4-����;5-ͣ��

    If MsgBox("��ȷ��" & IIf(bytType = 4, "����", "ͣ��") & "���Ϊ��" & strName & "���Ĳ���������", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo errH
    gstrSql = "Zl_ҽ�����Ӱಡ������_Edit(" & bytType & ",'" & strName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            
    With vsfPatiType
        If bytType = 4 Then
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlack
        Else
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub AdjustNum()
'�ϲ���,ֻ�����ڵĲſ��Ժϲ����ϲ��������ó���3��
    Dim i As Long, lngRow As Long, lngNum As Long, lngFirstNum As Long, lngSelRow As Long, lngRows As Long
    Dim strName As String, strTemp As String
    Dim blnCancel As Boolean
    Dim arrSql As Variant
    
    Dim objControl As CommandBarControl
        
    Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_RowProject)
    On Error GoTo errH
    With vsfPatiType
        strName = .TextMatrix(.Row, .ColIndex("���"))
    End With
    arrSql = Array()
    With vsfPatiTypeInfo
        If objControl.Caption = "ȡ���ϲ���" Then
            'ѡ���е����
            lngNum = .TextMatrix(.Row, .ColIndex("���"))
            For i = 1 To .Rows - 1
                blnCancel = False
                If .TextMatrix(i, .ColIndex("���")) = .TextMatrix(.Row, .ColIndex("���")) Then
                    blnCancel = True
                    If lngFirstNum = 0 Then lngFirstNum = i
                    .TextMatrix(i, .ColIndex("�ϲ�")) = ""
                End If
                '��ǰ����ȡ���ϲ���
                If blnCancel Then
                    If i > lngFirstNum And lngFirstNum <> 0 Then
                        lngNum = lngNum + 1
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = "Zl_ҽ�����Ӱಡ����Ŀ_Edit(4,'" & strName & "','" & _
                            .TextMatrix(i, .ColIndex("��Ŀ����")) & "',''," & lngNum & ")"
                    End If
                Else
                    'ѡ����ͬ��ŵ�һ��֮�����Ŷ�����������֮���кϲ��У��������ͬ
                    If i > lngFirstNum And lngFirstNum <> 0 Then
                        If lngRow <> .TextMatrix(i, .ColIndex("���")) Then
                            lngNum = lngNum + 1
                        End If
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = "Zl_ҽ�����Ӱಡ����Ŀ_Edit(4,'" & strName & "','" & _
                            .TextMatrix(i, .ColIndex("��Ŀ����")) & "',''," & lngNum & ")"
                    End If
                End If
                lngRow = .TextMatrix(i, .ColIndex("���"))
            Next
        Else
            lngSelRow = .Row
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = flexChecked Then
                    lngNum = lngNum + 1
                    If lngNum > 2 Then
                        MsgBox "�ϲ��в��ó���2�У�������ѡ��", vbInformation, Me.Caption
                        Exit Sub
                    End If
                    If lngRow = 0 Then
                        lngFirstNum = .TextMatrix(i, .ColIndex("���"))
                    Else
                        If i - lngRow <> 1 Then
                            MsgBox "�ϲ��б����������е����ݣ����飡", vbInformation, Me.Caption
                            Exit Sub
                        End If
                        If lngRows <> .TextMatrix(i, .ColIndex("��������")) Then
                            MsgBox "�ϲ��е�����������ͬ�����飡", vbInformation, Me.Caption
                            Exit Sub
                        End If
                    End If
                    lngRow = i
                    lngRows = .TextMatrix(i, .ColIndex("��������"))
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = "Zl_ҽ�����Ӱಡ����Ŀ_Edit(4,'" & strName & "','" & _
                        .TextMatrix(i, .ColIndex("��Ŀ����")) & "',''," & lngFirstNum & ")"
                End If
            Next
            If lngNum < 2 Then
                MsgBox "�ϲ��в��������������ݣ������Ƿ�ѡ��", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End With
    For i = 0 To UBound(arrSql)
        strTemp = arrSql(i)
        Call zlDatabase.ExecuteProcedure(strTemp, "�������")
    Next
    If objControl.Caption <> "ȡ���ϲ���" Then
        Call vsfPatiType_Click
        vsfPatiTypeInfo.Row = lngSelRow
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnStop As Boolean, blnPatiType As Boolean, blnPatiInfo As Boolean
    Dim blnPriv As Boolean
    
    With vsfPatiType
        If .Row > 0 Then
            blnPatiType = True
            blnStop = .Cell(flexcpForeColor, .Row, 1, .Row, .Cols - 1) = vbRed
        End If
    End With
    
    With vsfPatiTypeInfo
        If .Row > 0 Then
            blnPatiInfo = True
        End If
    End With
    blnPriv = CheckPrivs("��ɾ��")
    Select Case Control.ID
        Case conMenu_DocShift_File_Preview
            Control.Enabled = blnPriv And blnPatiType And Not blnStop
        Case conMenu_DocShift_Edit_New, conMenu_DocShift_Edit_Delete
            Control.Enabled = blnPriv And blnPatiType
        Case conMenu_DocShift_Edit_Modify
            Control.Enabled = blnPriv And blnPatiType And Not blnStop
        Case conMenu_DocShift_Edit_Reuse
            Control.Enabled = blnPriv And blnStop
        Case conMenu_DocShift_Edit_Stop
            Control.Enabled = blnPriv And Not blnStop
        Case conMenu_DocShift_Edit_NewProject
            Control.Enabled = blnPriv And Not blnStop
        Case conMenu_DocShift_Edit_ModifyProject, conMenu_DocShift_Edit_DeleteProject
            Control.Enabled = blnPriv And blnPatiInfo And Not blnStop
        Case conMenu_DocShift_Edit_RowProject
            '���ѡ�����Ǻϲ��У��򹤾�����ʾȡ���ϲ��У���֮������ʾ�ϲ���
            '��������caption������Ч
            Control.Enabled = False
            If blnPatiInfo And Not blnStop Then
                If vsfPatiTypeInfo.TextMatrix(vsfPatiTypeInfo.Row, vsfPatiTypeInfo.ColIndex("�ϲ�")) = "" Then
                    Control.Caption = "�ϲ���      "
                Else
                    Control.Caption = "ȡ���ϲ���"
                End If
            End If
            Control.Enabled = blnPriv And blnPatiInfo And Not blnStop
    End Select
End Sub

Private Function CheckPrivs(ByVal strPrivs As String) As Boolean
'����Ƿ���ָ����Ȩ��
    
    If InStr(";" & mstrPrivs & ";", ";" & strPrivs & ";") = 0 Then Exit Function
    CheckPrivs = True
End Function

Private Sub Form_Load()

    mstrPrivs = gstrPrivs
    Call InitCommandBar
    Call LoadData
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub PatiType(bytType As Byte, Optional ByVal strSName As String)
'bytType:1-������2-�޸�
    Dim rsTemp As ADODB.Recordset
    
    If frmDocShiftTypeEdit.ShowMe(bytType, strSName) Then
        Set rsTemp = rsPatiType(strSName)
        If rsTemp.RecordCount = 1 Then
            With vsfPatiType
                If bytType = 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                End If
                .TextMatrix(.Row, .ColIndex("���")) = rsTemp!���
                .TextMatrix(.Row, .ColIndex("��������")) = rsTemp!����
                .TextMatrix(.Row, .ColIndex("��ʼ����")) = rsTemp!��ʼ���� & ""
                .TextMatrix(.Row, .ColIndex("��ȡSQL")) = rsTemp!��ȡSQL & ""
            End With
            Call vsfPatiType_Click
        End If
    End If
End Sub

Private Sub PatiPrj(bytType As Byte, ByVal strSName As String, ByVal strPatiPrj As String)
'bytType:1-������2-�޸�
    Call frmDocShiftProEdit.ShowMe(bytType, strSName, strPatiPrj)
End Sub

Public Sub RefreshPrj(ByVal bytType As Byte)
'bytType:1-������2-�޸�
'������Ŀ���Ա����������������ӽ�������ɹ������涼��Ҫˢ��
    Dim lngRow As Long
    
    lngRow = vsfPatiTypeInfo.Row
    Call vsfPatiType_Click
    If bytType = 1 Then
        vsfPatiTypeInfo.Row = vsfPatiTypeInfo.Rows - 1
    Else
        vsfPatiTypeInfo.Row = lngRow
    End If
End Sub

Private Sub LoadData()
    Dim rsTemp As ADODB.Recordset
    
    Set rsTemp = GetPatiType
    With vsfPatiType
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("˳��")) = Val(rsTemp!˳��)
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("���")) = rsTemp!���
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("��������")) = rsTemp!����
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("��ʼ����")) = rsTemp!��ʼ���� & ""
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("��ȡSQL")) = rsTemp!��ȡSQL & ""
            If rsTemp!�Ƿ�ͣ�� = 1 Then
                .Cell(flexcpForeColor, rsTemp.AbsolutePosition, 0, rsTemp.AbsolutePosition, .Cols - 1) = vbRed
            End If
            rsTemp.MoveNext
        Loop
        If .Rows > 1 Then
            .Row = 1
            Call vsfPatiType_Click
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Function GetPatiType() As ADODB.Recordset
'��ȡ��������
    
    gstrSql = "Select ˳��, ���, ����, ��ʼ����, ��ȡsql, �Ƿ�ͣ�� From ҽ�����Ӱಡ������ Order By ˳��"
    Set GetPatiType = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ��������")
End Function

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    On Error Resume Next
    
    If Not cbsMain(2).Visible Then
        lngTop = 500
    End If
    If stbThis.Visible Then
        lngHeight = stbThis.Height
    End If
    
    lblPatiType.Move 120, 1000 - lngTop
    vsfPatiType.Move 120, lblPatiType.Top + lblPatiType.Height + 100, 3075, Me.ScaleHeight - 1000 - lngHeight + lngTop - lblPatiType.Height - 100
    lblPatiTypeInfo.Move vsfPatiType.Left + vsfPatiType.Width + 150, lblPatiType.Top
    vsfPatiTypeInfo.Move lblPatiTypeInfo.Left, vsfPatiType.Top, Me.ScaleWidth - vsfPatiType.Width - 500, vsfPatiType.Height
End Sub

Private Sub vsfPatiType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfPatiType
        If NewRow = 1 Then
            If .Rows = 2 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
        End If
    End With
End Sub

Private Sub vsfPatiType_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsfPatiType
        If Not (.Col = .ColIndex("����") Or .Col = .ColIndex("����") Or .Col = .ColIndex("ѡ��")) Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfPatiType_Click()
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim lngNum As Long, lngRow As Long, i As Long
    Dim blnAdd As Boolean
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
    With vsfPatiType
        If vsfPatiType.Row <= 0 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_Modify)
        If objControl.Enabled Then
            If .Col = .ColIndex("����") Then
                If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                    lngRow = .Row - 1
                End If
            ElseIf .Col = .ColIndex("����") Then
                If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                    lngRow = .Row + 1
                End If
            End If
        End If
        If lngRow <> 0 Then
            '����˳�򲻱䣬���Դ�1��ʼ
            For i = 1 To .ColIndex("��������")
                strTemp = .TextMatrix(.Row, i)
                .TextMatrix(.Row, i) = .TextMatrix(lngRow, i)
                .TextMatrix(lngRow, i) = strTemp
            Next
            gstrSql = "Zl_ҽ�����Ӱಡ������_Edit(6,'" & .TextMatrix(.Row, .ColIndex("���")) & "','','','',''," & _
                Val(.TextMatrix(.Row, .ColIndex("˳��"))) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, "����˳��")
            gstrSql = "Zl_ҽ�����Ӱಡ������_Edit(6,'" & .TextMatrix(lngRow, .ColIndex("���")) & "','','','',''," & _
                Val(.TextMatrix(lngRow, .ColIndex("˳��"))) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, "����˳��")
            .Row = lngRow
        End If
    End With
    
    Set rsTemp = GetPatiTypeInfo(vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("���")))
    With vsfPatiTypeInfo
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            lngRow = rsTemp.AbsolutePosition
            If rsTemp!��� = lngNum Then
                '�����ͬ�����س�����������ʾΪ������������
                If blnAdd Then
                    If lngRow = .Rows - 1 Then
                        .TextMatrix(lngRow, .ColIndex("�ϲ�")) = "��"
                    End If
                Else
                    If lngRow = .Rows - 1 Then
                        .TextMatrix(lngRow - 1, .ColIndex("�ϲ�")) = "��"
                        .TextMatrix(lngRow, .ColIndex("�ϲ�")) = "��"
                    Else
                        blnAdd = True
                        .TextMatrix(lngRow - 1, .ColIndex("�ϲ�")) = "��"
                        .TextMatrix(lngRow, .ColIndex("�ϲ�")) = "��"
                    End If
                End If
            Else
                If blnAdd Then
                    blnAdd = False
                    .TextMatrix(lngRow - 1, .ColIndex("�ϲ�")) = "��"
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("��Ŀ����")) = rsTemp!��Ŀ����
            .TextMatrix(lngRow, .ColIndex("���")) = Val(rsTemp!��� & "")
            .TextMatrix(lngRow, .ColIndex("��Ŀ���")) = rsTemp!��Ŀ��� & ""
            strTemp = rsTemp!������ʽ & ""
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            .TextMatrix(lngRow, .ColIndex("������ʽ")) = strTemp
            strTemp = rsTemp!�������� & ""
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            .TextMatrix(lngRow, .ColIndex("��������")) = strTemp
            .TextMatrix(lngRow, .ColIndex("�����ʽ")) = rsTemp!�����ʽ & ""
            .TextMatrix(lngRow, .ColIndex("����ֵ��")) = rsTemp!����ֵ�� & ""
            .TextMatrix(lngRow, .ColIndex("��������")) = rsTemp!�������� & ""
            strTemp = rsTemp!��ȡ��Դ & ""
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            .TextMatrix(lngRow, .ColIndex("��ȡ��Դ")) = strTemp
            .TextMatrix(lngRow, .ColIndex("��ȡ����")) = rsTemp!��ȡ���� & ""
            .TextMatrix(lngRow, .ColIndex("��ȡSQL")) = rsTemp!��ȡSQL & ""
            .TextMatrix(lngRow, .ColIndex("��������")) = rsTemp!�������� & ""
            .TextMatrix(lngRow, .ColIndex("�Ƿ�ֻ��")) = rsTemp!�Ƿ�ֻ�� & ""
            .TextMatrix(lngRow, .ColIndex("����������")) = rsTemp!���������� & ""
            '��������ͬ���������Ҫչ�ֳ���
            lngNum = Val(rsTemp!��� & "")
            rsTemp.MoveNext
        Loop
        If .Rows > 1 Then .Row = 1
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = imgPublic.Icons

    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_DocShift_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Preview, "Ԥ��(&V)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_PatiTypePopup, "��������(&E)", -1, False)
    objMenu.ID = conMenu_DocShift_PatiTypePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_New, "����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Delete, "ɾ��(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Reuse, "����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Stop, "ͣ��(&S)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_PatiProjectPopup, "������Ŀ(&E)", -1, False)
    objMenu.ID = conMenu_DocShift_PatiProjectPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_NewProject, "����(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_ModifyProject, "�޸�(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_DeleteProject, "ɾ��(&S)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_DocShift_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_DocShift_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
            objControl.Checked = True
        End With
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_StatusBar, "״̬��(&S)")
        objControl.Checked = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_DocShift_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_DocShift_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_DocShift_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_DocShift_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_New, "��������"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Modify, "�޸�����")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Delete, "ɾ������")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Reuse, "����(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Stop, "ͣ��(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_NewProject, "������Ŀ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_ModifyProject, "�޸���Ŀ")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_DeleteProject, "ɾ����Ŀ")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_RowProject, "�ϲ���      ")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Exit, "�˳�"): objControl.BeginGroup = True
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next

    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
'
'        .Add FCONTROL, vbKeyA, conMenu_DocShift_Edit_New '����
'        .Add FCONTROL, vbKeyM, conMenu_DocShift_Edit_Modify '�޸�
'        .Add 0, vbKeyDelete, conMenu_DocShift_Edit_Delete 'ɾ��
    End With
End Sub

Private Sub vsfPatiType_DblClick()
    Dim objControl As CommandBarControl
    
    With vsfPatiType
        If .MouseRow < 1 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_Modify)
        If objControl.Enabled = False Then Exit Sub
        Call PatiType(2, .TextMatrix(.Row, .ColIndex("���")))
    End With
End Sub

Private Sub vsfPatiTypeInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim blnCheck As Boolean
    Dim lngNum As Long, i As Long
    
    With vsfPatiTypeInfo
        If Col = .ColIndex("ѡ��") Then
            blnCheck = .Cell(flexcpChecked, Row, Col) = flexChecked
            lngNum = .TextMatrix(Row, .ColIndex("���"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���")) = lngNum Then
                    .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = IIf(blnCheck, flexChecked, flexUnchecked)
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPatiTypeInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long, lngNum As Long
    Dim blnUp As Boolean, blnDown As Boolean
    Dim lngBegin As Long, lngEnd As Long
    
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfPatiTypeInfo
        lngNum = Val(.TextMatrix(NewRow, .ColIndex("���")))
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("���"))) = lngNum Then
                '��¼�ϲ��еĵ�һ�е��к�
                If lngBegin = 0 Then lngBegin = i
                '��¼�ϲ��е����һ�е��к�
                lngEnd = i
            End If
        Next
        '����Ǻϲ��е����ݣ�����ͬ��ŵĵ�һ�����ж��ܷ����ƣ�����ͬ��ŵ����һ���ж��ܷ�����
        If lngBegin = 1 Then
            If .Rows > 2 Then
                blnUp = False
            End If
        Else
            If lngBegin = .Rows - 1 Then
                blnUp = True
            Else
                blnUp = True
            End If
        End If
        If lngEnd = 1 Then
            If .Rows > 2 Then
                blnDown = True
            End If
        Else
            If lngEnd = .Rows - 1 Then
                blnDown = False
            Else
                blnDown = True
            End If
        End If
        .Cell(flexcpPicture, NewRow, .ColIndex("����")) = IIf(blnUp, imgList.ListImages("Up").Picture, "")
        .Cell(flexcpPicture, NewRow, .ColIndex("����")) = IIf(blnDown, imgList.ListImages("Down").Picture, "")
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
        End If
    End With
End Sub

Private Sub vsfPatiTypeInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPatiTypeInfo
        If Not (.Col = .ColIndex("����") Or .Col = .ColIndex("����") Or .Col = .ColIndex("ѡ��")) Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfPatiTypeInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfPatiTypeInfo.ColIndex("��Ŀ����") Then Cancel = True
End Sub

Private Sub vsfPatiTypeInfo_Click()
    Dim lngRow As Long, i As Long, j As Long, lngChangeRow As Long, lngNum As Long, lngChangeNum As Long, m As Long
    Dim strPrj As String, strTemp As String
    Dim strName As String
    Dim blnUp As Boolean, blnDown As Boolean, blnFirst As Boolean, blnAddRow As Boolean
    Dim arrSql As Variant
    Dim objControl As CommandBarControl
        
    If vsfPatiTypeInfo.Row < 1 Then Exit Sub
    strName = vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("���"))
    With vsfPatiTypeInfo
        If .Row < 1 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_ModifyProject)
        If objControl.Enabled Then
            If .Col = .ColIndex("����") Then
                '�ҳ�Ҫ�������е��кţ������漰�ϲ��У�������ȡ�ϲ��е�һ�е���һ�У�����ȡ�ϲ������һ�е���һ��
                If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, .ColIndex("���")) = .TextMatrix(.Row, .ColIndex("���")) Then
                            lngRow = i - 1
                            Exit For
                        End If
                    Next
                    blnUp = True
                End If
            ElseIf .Col = .ColIndex("����") Then
                If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, .ColIndex("���")) = .TextMatrix(.Row, .ColIndex("���")) Then
                            lngRow = i + 1
                        End If
                    Next
                    blnDown = True
                End If
            End If
        End If
        If blnUp = False And blnDown = False Then Exit Sub
        
        lngChangeNum = .TextMatrix(lngRow, .ColIndex("���")) 'Ҫ���ƻ������ƺ�����
        lngNum = .TextMatrix(.Row, .ColIndex("���")) '��ǰѡ���е����
        strPrj = .TextMatrix(.Row, .ColIndex("��Ŀ����")) 'ѡ���е�����
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = lngChangeNum Then
                If blnUp Then
                    '����ȡ������ͬ��ŵĵ�һ��
                    lngChangeRow = i - 1
                    Exit For
                Else
                    '����ȡ������ͬ��ŵ����һ��
                    lngChangeRow = i
                End If
            End If
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = lngChangeNum Then
                '��Ҫ�ƶ�����е���ŵ���Ϊ-1
                .TextMatrix(i, .ColIndex("���")) = -1
            End If
        Next
        '���ڲ�����ʱ�����������ӣ�������������˶�����
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = lngNum Then
                m = m + 1
            End If
        Next
        '�ò����еķ�����ʵ�����ƺ�����
        For i = 1 To .Rows - 1 + m
            If .TextMatrix(i, .ColIndex("���")) = lngNum Then
                '����������㷨�ó����к�
                lngChangeRow = lngChangeRow + 1
                .AddItem "", lngChangeRow
                If blnFirst = False Then
                    blnFirst = True
                    '����������ѡ���е�����ʱ����֤���ݵ���ȷ�ԣ�ѡ����Ӧ�����ƶ�һ��
                    If lngChangeRow < .Row Then
                        .Row = .Row + 1
                        blnAddRow = True
                    End If
                End If
                For j = .ColIndex("��Ŀ����") To .ColIndex("����������")
                    .TextMatrix(lngChangeRow, j) = .TextMatrix(IIf(blnAddRow, i + 1, i), j)
                Next
                .TextMatrix(lngChangeRow, .ColIndex("�ϲ�")) = .TextMatrix(IIf(blnAddRow, i + 1, i), .ColIndex("�ϲ�"))
                .TextMatrix(lngChangeRow, .ColIndex("���")) = lngChangeNum
                .TextMatrix(IIf(blnAddRow, i + 1, i), .ColIndex("���")) = 0
            End If
        Next
        '��������ȷ�����
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, .ColIndex("���")) = -1 Then
                .TextMatrix(i, .ColIndex("���")) = lngNum
            ElseIf .TextMatrix(i, .ColIndex("���")) = 0 Then
                .RemoveItem i
            End If
        Next
        arrSql = Array()
        For i = 1 To .Rows - 1
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "Zl_ҽ�����Ӱಡ����Ŀ_Edit(4,'" & strName & "','" & .TextMatrix(i, .ColIndex("��Ŀ����")) & "',''," & _
            .TextMatrix(i, .ColIndex("���")) & ")"
            If .TextMatrix(i, .ColIndex("��Ŀ����")) = strPrj Then
                .Row = i
            End If
        Next
        On Error GoTo errH
        For i = 0 To UBound(arrSql)
            strTemp = arrSql(i)
            Call zlDatabase.ExecuteProcedure(strTemp, "�������")
        Next
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub vsfPatiTypeInfo_DblClick()
    Dim objControl As CommandBarControl
    
    With vsfPatiTypeInfo
        If .MouseRow < 1 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_ModifyProject)
        If objControl.Enabled = False Then Exit Sub
        If .Col < .ColIndex("��Ŀ����") Then Exit Sub
        Call PatiPrj(2, vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("���")), _
            .TextMatrix(.Row, .ColIndex("��Ŀ����")))
    End With
End Sub

