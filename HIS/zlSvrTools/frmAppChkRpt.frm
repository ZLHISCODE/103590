VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppChkRpt 
   Caption         =   "��������"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13860
   Icon            =   "frmAppChkRpt.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   13860
   StartUpPosition =   2  '��Ļ����
   Tag             =   "17500"
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   0
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   270
      Width           =   2205
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   1
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1500
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   2
      Left            =   12240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   270
      Width           =   1200
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "�����Excel"
      Height          =   350
      Left            =   6000
      TabIndex        =   3
      Top             =   7560
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfResult 
      Height          =   6255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   10695
      _cx             =   18865
      _cy             =   11033
      Appearance      =   3
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   9390
      TabIndex        =   1
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "����"
      Height          =   350
      Left            =   8205
      TabIndex        =   0
      Top             =   7680
      Width           =   1100
   End
   Begin VB.Label lblRsFilter 
      Caption         =   "Label1"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   7560
      Width           =   5535
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "ϵͳ"
      Height          =   180
      Index           =   0
      Left            =   5760
      TabIndex        =   9
      Top             =   315
      Width           =   360
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   8520
      TabIndex        =   8
      Top             =   315
      Width           =   360
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "���س̶�"
      Height          =   180
      Index           =   2
      Left            =   11400
      TabIndex        =   7
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmAppChkRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_COL = ",300,4;���,500,4;ϵͳ,2000,1;����,1500,1;������,2450,1;��������,6300,1;����˵��,3000,1;���س̶�,930,4;����SQL,0,4"
Private mrsProData As New ADODB.Recordset
Private mrsDataFromFile As New ADODB.Recordset
Private mstrSysModul As String
Public Enum enuResult
    Col_ѡ�� = 0
    Col_��� = 1
    Col_ϵͳ = 2
    Col_���� = 3
    Col_������ = 4
    Col_�������� = 5
    Col_����˵�� = 6
    Col_���س̶� = 7
    Col_����SQL = 8
End Enum

Private mblnFirst As Boolean
Private mstrPath As String

Private Sub cboFilter_Click(Index As Integer)
    Dim strFilter As String
    
    If mblnFirst = False Then Exit Sub
    
    If cboFilter(0).Text = "����ϵͳ" Then
        strFilter = ""
    Else
        strFilter = "ϵͳ����='" & cboFilter(0).Text & "'"
    End If
    
    If cboFilter(1).Text = "��������" Then
        strFilter = strFilter
    Else
        strFilter = IIf(strFilter = "", "���='" & cboFilter(1).Text & "'", strFilter & " and ���='" & cboFilter(1).Text & "'")
    End If
    
    If cboFilter(2).Text = "���г̶�" Then
        strFilter = strFilter
    Else
        strFilter = IIf(strFilter = "", "���س̶�='" & cboFilter(2).Text & "'", strFilter & " and ���س̶�='" & cboFilter(2).Text & "'")
    End If
    
    Call AddvsfData(strFilter)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    mblnFirst = False
    Call InitTable(vsfResult, MSTR_COL)
    Call InivsfData
    mblnFirst = True
End Sub

Public Function ShowMe(ByVal strPath As String, ByVal rsProData As ADODB.Recordset, ByVal rsDataFromFile As ADODB.Recordset) As Boolean

    Set mrsProData = rsProData
    Set mrsDataFromFile = rsDataFromFile
    
    mstrPath = strPath & "\Log\��־����\zlObjCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".Log"
    Me.Show 1
End Function

Private Sub InivsfData()
'���ܣ���һ�ν�����ʾ����
    Dim i As Long
    Dim strSys As String
    Dim strType2 As String
    Dim strSer As String
    
    With vsfResult
        strSys = "����ϵͳ"
        strType2 = "��������"
        strSer = "���г̶�"
        cboFilter(0).addItem "����ϵͳ"
        cboFilter(1).addItem "��������"
        cboFilter(2).addItem "���г̶�"
        cboFilter(2).addItem "����"
        cboFilter(2).addItem "����"
        cboFilter(2).addItem "��΢"
        .rowHeight(0) = 500
        .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
        .Rows = .Rows - 1
        Call AddvsfData
        
        For i = 1 To .Rows - 1
            If InStr(strSys, .TextMatrix(i, Col_ϵͳ)) = 0 Then
                strSys = strSys & "|" & .TextMatrix(i, Col_ϵͳ)
                cboFilter(0).addItem .TextMatrix(i, Col_ϵͳ)
            End If
            
            If InStr(strType2, .TextMatrix(i, Col_����)) = 0 Then
                strType2 = strType2 & "|" & .TextMatrix(i, Col_����)
                cboFilter(1).addItem .TextMatrix(i, Col_����)
            End If
        Next
    End With
    
    cboFilter(0).ListIndex = 0
    cboFilter(1).ListIndex = 0
    cboFilter(2).ListIndex = 0
End Sub

Private Sub AddvsfData(Optional ByVal strFilter As String)
'���ܣ����������󵽱����
    Dim i As Long
    
    With vsfResult
        .Rows = 1
        .ColHidden(Col_����SQL) = True
        mrsProData.Filter = strFilter
        Do While Not mrsProData.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, Col_���) = .Rows - 1
            .TextMatrix(.Rows - 1, Col_ϵͳ) = mrsProData!ϵͳ����
            .TextMatrix(.Rows - 1, Col_����) = mrsProData!���
            .TextMatrix(.Rows - 1, Col_������) = mrsProData!������
            .TextMatrix(.Rows - 1, Col_��������) = mrsProData!��������
            .TextMatrix(.Rows - 1, Col_����˵��) = mrsProData!����˵��
            .TextMatrix(.Rows - 1, Col_���س̶�) = mrsProData!���س̶�
            .TextMatrix(.Rows - 1, Col_����SQL) = mrsProData!����SQL
            .rowHeight(.Rows - 1) = 500
            If .TextMatrix(.Rows - 1, Col_���س̶�) = "��΢" Then
                .Cell(flexcpBackColor, .Rows - 1, Col_���س̶�) = RGB(238, 230, 133)
            ElseIf .TextMatrix(.Rows - 1, Col_���س̶�) = "����" Then
                .Cell(flexcpBackColor, .Rows - 1, Col_���س̶�) = RGB(238, 201, 0)
            ElseIf .TextMatrix(.Rows - 1, Col_���س̶�) = "����" Then
                .Cell(flexcpBackColor, .Rows - 1, Col_���س̶�) = RGB(238, 154, 0)
            End If
            If InStr(.TextMatrix(.Rows - 1, Col_����˵��), "�˹�") > 0 Then
                .TextMatrix(.Rows - 1, Col_ѡ��) = ""
            Else
                .Cell(flexcpChecked, .Rows - 1, Col_ѡ��) = flexUnchecked
            End If
            mrsProData.MoveNext
        Loop
        .Cell(flexcpAlignment, 0, 0, .Rows - 1) = 4
        If .Rows > 1 Then
            .Row = 1
            Call .ShowCell(1, 1)
        End If
    End With
    lblRsFilter.Caption = "���������" & mrsProData.RecordCount & "�����⡣"
End Sub

Private Sub cmdModify_Click()
'���ܣ�������ѡ�Ķ�������
    Dim i As Long
    Dim j As Long
    Dim varTemp As Variant
    Dim strErr As String
    Dim strTemp As String
    Dim strSQL As String
    Dim blnModify As Boolean
    Dim blnFalse As Boolean
    Dim cnChoose As ADODB.Connection
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
            
    With vsfResult
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                Call ShowFlash("���ڽ��ж�������ݵ����������Ժ�")
                If .TextMatrix(i, Col_ϵͳ) = "������������" Then
                    If gcnTools Is Nothing Then
                        Set gcnTools = GetConnection("ZLTOOLS")
                    End If
                    Set cnChoose = gcnTools
                Else
                    Set cnChoose = gcnOracle
                End If
                blnFalse = True
                varTemp = Split(.TextMatrix(i, Col_����SQL), "{JM|SQL�ָ���}" & vbNewLine)
                For j = 0 To UBound(varTemp)
                    strSQL = varTemp(j)
                    If strSQL <> "" Then
                        On Error Resume Next
                        cnChoose.Execute strSQL
                        If err.Number <> 0 Then
                            If strSQL Like "INSERT INTO ZLPARAMETERS*" Then
                                strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                                Set rsTemp = SetSelectRecordset(strSQL, strTemp, Split(strTemp, ","), "ZLPARAMETERS")
                                If InStr(rsTemp!ģ��, "NULL") = 0 And InStr(rsTemp!ϵͳ, "NULL") = 0 Then
                                    If InStr(mstrSysModul, rsTemp!ϵͳ & "&" & rsTemp!ģ��) = 0 Then
                                        mrsDataFromFile.Filter = "���='����'"
                                        Set rsData = CopyNewRec(mrsDataFromFile)
                                        mstrSysModul = mstrSysModul & "|" & rsTemp!ϵͳ & "&" & rsTemp!ģ��
                                        strSQL = "Update Zlparameters Set ������ = -1 * ������ Where ϵͳ =" & rsTemp!ϵͳ & " And ģ�� = " & rsTemp!ģ��
                                        cnChoose.Execute strSQL
                                        rsData.Filter = "���='����' and ����=" & rsTemp!ģ�� & " and ϵͳ���=" & rsTemp!ϵͳ
                                        Do While Not rsData.EOF
                                            mrsDataFromFile.Filter = "���='����' and ����=" & rsTemp!ģ�� & " and ϵͳ���=" & rsTemp!ϵͳ & " and ������='" & rsData!������ & "'"
                                            If mrsDataFromFile.RecordCount > 0 Then
                                                strSQL = "Update Zlparameters Set ������ = " & rsTemp!������ & " Where ϵͳ =" & rsTemp!ϵͳ & " And ģ�� = " & rsTemp!ģ�� & " and ������='" & rsData!������ & "'"
                                                cnChoose.Execute strSQL
                                            End If
                                            rsData.MoveNext
                                        Loop
                                        cnChoose.Execute varTemp(j)
'                                        strSQL = "Update Zlparameters Set ������ = -1 * ������ Where ϵͳ =" & rsTemp!ϵͳ & " And ģ�� = " & rsTemp!ģ��
'                                        cnChoose.Execute strSQL
                                    End If
                                Else
                                    blnFalse = False
                                    strErr = IIf(strErr = "", "����ʧ�ܵ�SQL��" & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf, strErr & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf)
                                End If
                            Else
                                blnFalse = False
                                strErr = IIf(strErr = "", "����ʧ�ܵ�SQL��" & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf, strErr & vbCrLf & varTemp(j) & ";" & vbCrLf & "ԭ��:" & err.Description & vbCrLf)
                            End If
                        End If
                    End If
                Next
                blnModify = True
                If blnFalse Then
                    .Cell(flexcpData, i, 0) = 1
                Else
                    .Cell(flexcpData, i, 0) = 0
                End If
            End If
        Next
        If blnModify = False Then
            MsgBox "δ��ѡ���Զ����������ݣ�"
            Exit Sub
        End If
        Call ShowFlash("")
        If strErr <> "" Then
            On Error Resume Next
            Call WriteErrorLog(strErr)
            If err.Number = 0 Then
                MsgBox "������ɣ��в�������δ�ɹ��������������" & mstrPath
            Else
                MsgBox "������ɣ�������־��¼ʧ�ܣ������Ǹ���־�ļ�(" & mstrPath & ")�Ѵ򿪣����飡"
            End If
            err.Clear: On Error GoTo 0
        Else
            MsgBox "������ɣ�"
        End If
    End With
    Call AfterModify
End Sub

Private Sub AfterModify()
'������ɺ�����ˢ�½�������
    Dim i As Long
    Dim strFilter As String
    Dim lngSelRow As Long
    
    lblRsFilter.Caption = "��������ˢ�½���......"
    With vsfResult
        lngSelRow = .Row
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                strFilter = "��������='" & .TextMatrix(i, Col_��������) & "' and ������='" & .TextMatrix(i, Col_������) & "' and ���='" & .TextMatrix(i, Col_����) & "'"
                Call RecDelete(mrsProData, strFilter)
            End If
        Next
        Call cboFilter_Click(0)
        If .Rows > 1 Then
            If .Rows > lngSelRow Then
                .Row = lngSelRow
                Call .ShowCell(lngSelRow, 1)
            Else
                .Row = .Rows - 1
                Call .ShowCell(.Rows - 1, 1)
            End If
        End If
    End With
    Call vsfResult_AfterEdit(1, 0)
End Sub

Private Sub WriteErrorLog(ByVal strErr As String)
    Dim objFile As Object
    Dim objStream As TextStream
    Dim strPath As String
        
    Set objFile = CreateObject("Scripting.FileSystemObject")
    If objFile.FileExists(mstrPath) = False Then objFile.CreateTextFile mstrPath
    Set objStream = objFile.OpenTextFile(mstrPath)

    Open mstrPath For Append Shared As #1
    Print #1, strErr
    Close #1
End Sub

Private Sub cmdOut_Click()
    
    Call OutExcel
End Sub

Private Sub OutExcel()
'���ܣ���vsf����������Excel��
    Dim strPath As String
    Dim spShell, spFolder, spFolderItem, spPath As String
    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = 0

    If IsInstallExcel Then
        With vsfResult
            If .Rows < 2 Then
                MsgBox "�����û�����ݣ��޷�������ݣ����飡"
                Exit Sub
            Else
                Set spShell = CreateObject("Shell.Application")
                Set spFolder = spShell.BrowseForFolder(WINDOW_HANDLE, "ѡ��Ŀ¼:", NO_OPTIONS)
                If spFolder Is Nothing Then
                    Exit Sub
                Else
                    Set spFolderItem = spFolder.Self
                    spPath = spFolderItem.Path
                    .SaveGrid Replace(spPath & "\zlObjectCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".xls", "\\", "\"), flexFileExcel, True
                    .BackColorSel = &H8000000D
                     MsgBox "����ɹ���������ѱ�����" & Replace(spPath & "\zlObjectCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".xls", "\\", "\")
                     Exit Sub
                End If
            End If
        End With
    End If
errH:
    MsgBox "��ѡ·���ĸ��ļ����ڴ�״̬����ѡ·������"
End Sub

Private Sub Form_Resize()

    If ScaleHeight < 2000 Then Exit Sub
    
    With vsfResult
        .Top = ScaleTop + 600
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = ScaleHeight - cmdModify.Height - 900
        .ColWidth(Col_����) = 1500 + 0.05 * (Me.Width - Me.Tag)
        .ColWidth(Col_������) = 2450 + 0.25 * (Me.Width - Me.Tag)
        .ColWidth(Col_��������) = 6300 + 0.3 * (Me.Width - Me.Tag)
        .ColWidth(Col_����˵��) = 3000 + 0.4 * (Me.Width - Me.Tag)
    End With
    
    cmdClose.Top = vsfResult.Top + vsfResult.Height + 150
    cmdClose.Left = ScaleWidth - cmdClose.Width - 300
    
    cmdModify.Top = cmdClose.Top
    cmdModify.Left = cmdClose.Left - cmdModify.Width - 500
    
    cmdOut.Top = cmdClose.Top
    cmdOut.Left = cmdModify.Left - cmdOut.Width - 500
    
    lblRsFilter.Top = cmdOut.Top + 150
    lblRsFilter.Left = 300
    
    cboFilter(2).Top = 200
    cboFilter(2).Left = ScaleWidth - cboFilter(2).Width - 300
    lblFilter(2).Top = 250
    lblFilter(2).Left = cboFilter(2).Left - lblFilter(2).Width - 150
    
    cboFilter(1).Top = 200
    cboFilter(1).Left = lblFilter(2).Left - cboFilter(1).Width - 300
    lblFilter(1).Top = 250
    lblFilter(1).Left = cboFilter(1).Left - lblFilter(1).Width - 150
    
    cboFilter(0).Top = 200
    cboFilter(0).Left = lblFilter(1).Left - cboFilter(0).Width - 300
    lblFilter(0).Top = 250
    lblFilter(0).Left = cboFilter(0).Left - lblFilter(0).Width - 150

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call ReleaseMe
End Sub

Private Sub vsfResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strFilter As String
    Dim strTemp As String
    
    With vsfResult
        If Col = Col_ѡ�� Then
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked Then
                    .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked Then
                            .Cell(flexcpChecked, i, Col_ѡ��) = flexChecked
                        End If
                    Next
                Else
                    .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, Col_ѡ��) = flexChecked Then
                            .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked
                        End If
                    Next
                End If
            Else
                If .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked Then
                    .Cell(flexcpChecked, 0, Col_ѡ��) = flexUnchecked
                End If
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, Col_ѡ��) = flexUnchecked Then
                        Exit For
                    Else
                        If i = .Rows - 1 Then
                            .Cell(flexcpChecked, 0, Col_ѡ��) = flexChecked
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Public Function IsInstallExcel() As Boolean
'���ܣ��жϱ�����װ��EXCELû��
'���أ����򷵻�True
    Dim arrTemp  As Object
    
    On Error GoTo errH
    Set arrTemp = CreateObject("Excel.Application") '��һ��EXCEL����
    Set arrTemp = Nothing
    IsInstallExcel = True
    Exit Function
errH:
    Set arrTemp = Nothing
    IsInstallExcel = False
    MsgBox "�õ�����û�а�װEXCEL�������飡", vbCritical, GSTR_APPNAME
End Function

Private Sub vsfResult_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsfResult_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strTip As String
    
    With vsfResult
        If .MouseRow <> -1 And .MouseRow <> 0 And .MouseCol = Col_����˵�� Then
            If .TextMatrix(.MouseRow, Col_����SQL) <> "" Then
                strTip = "����SQL:" & vbNewLine & Replace(.TextMatrix(.MouseRow, Col_����SQL), "{JM|SQL�ָ���}", "")
                Call ShowTipInfo(.hwnd, strTip, True)
            Else
                Call ShowTipInfo(.hwnd, "")
            End If
        Else
            Call ShowTipInfo(.hwnd, "")
        End If
    End With
    
End Sub
