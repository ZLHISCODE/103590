VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckPartEdit 
   BorderStyle     =   0  'None
   Caption         =   "��鲿λ�༭"
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cbo�����Ա� 
      Height          =   300
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1920
      Width           =   2160
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   3855
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   2565
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "���Ϊ���ӿ�ѡ����(&A)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   2
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":0000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3630
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "���Ϊ���ÿ�ѡ����(&P)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":014A
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3300
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "���Ϊ�����������(&B)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":0294
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2970
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "ɾ����ǰ����(&D)      "
      Enabled         =   0   'False
      Height          =   350
      Index           =   3
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":03DE
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4155
      Width           =   2160
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -390
      TabIndex        =   15
      Top             =   1020
      Width           =   7320
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1755
      MaxLength       =   60
      TabIndex        =   3
      Top             =   135
      Width           =   1920
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   555
      MaxLength       =   4
      TabIndex        =   1
      Top             =   135
      Width           =   525
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   4410
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   135
      Width           =   1620
   End
   Begin VB.TextBox txt��ע 
      Height          =   300
      Left            =   555
      MaxLength       =   60
      TabIndex        =   7
      Top             =   555
      Width           =   5460
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg���� 
      Height          =   3090
      Left            =   135
      TabIndex        =   9
      Top             =   1410
      Width           =   3555
      _cx             =   6271
      _cy             =   5450
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin MSComctlLib.ImageList imgList 
         Left            =   2985
         Top             =   1860
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheckPartEdit.frx":0528
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheckPartEdit.frx":0AC2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lbl�����Ա� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ա�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3840
      TabIndex        =   19
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCheckPartEdit.frx":105C
      ForeColor       =   &H00008000&
      Height          =   1980
      Left            =   345
      TabIndex        =   16
      Top             =   4740
      Width           =   5580
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   15
      Picture         =   "frmCheckPartEdit.frx":1202
      Top             =   4710
      Width           =   240
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3855
      TabIndex        =   10
      Top             =   2325
      Width           =   720
   End
   Begin VB.Label lbl��֯ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��鷽����������֯:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   1170
      Width           =   1710
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1335
      TabIndex        =   2
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3975
      TabIndex        =   4
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lbl��ע 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ע"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   615
      Width           =   360
   End
End
Attribute VB_Name = "frmCheckPartEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ���� = 0
    ��Ӱ = 2
End Enum

Private mstrKind As String          '��ǰ����
Private mstrPart As String          '��ǰ��λ
Private mblnPACSInterface As Boolean        '����Ӱ����Ϣϵͳ�ӿ�
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub FormatList(Optional strMode As String)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������strMode-������
    Dim aryItem() As String, strItems As String, strTemp As String
    Dim aryChild() As String, lngChild As Long
    With Me.vfg����
        .Redraw = flexRDNone
        .Clear
        .Rows = 1: .FixedRows = 1: .Cols = 3: .FixedCols = 0
        .TextMatrix(0, mCol.����) = "��鷽��": .ColWidth(mCol.����) = 280: .FixedAlignment(mCol.����) = flexAlignCenterCenter
        .TextMatrix(0, mCol.���� + 1) = "��鷽��": .ColWidth(mCol.���� + 1) = 2500
        .TextMatrix(0, mCol.��Ӱ) = "��Ӱ"
        .MergeCells = flexMergeFree: .MergeRow(0) = True
        If strMode = "" Then .Redraw = flexRDDirect: Exit Sub
        
        strItems = ""
        strTemp = ""
        If InStr(1, strMode, vbTab) > 0 Then strMode = Mid(strMode, 1, InStr(1, strMode, vbTab) - 1) & ";" & Mid(strMode, InStr(1, strMode, vbTab))
        For lngCount = 1 To Len(strMode)
            If Mid(strMode, lngCount, 1) = vbTab And lngCount <> 2 Then
                 If Mid(strTemp, Len(strTemp), 1) <> ";" Then strTemp = strTemp & ";"
            End If
            strTemp = strTemp & Mid(strMode, lngCount, 1)
        Next
        strMode = strTemp
        
        aryItem() = IIf(Mid(strMode, 1, 1) = ";", Split(Mid(strMode, 2), ";"), Split(strMode, ";"))
        For lngCount = 0 To UBound(aryItem)
            strTemp = aryItem(lngCount)
            If InStr(1, aryItem(lngCount), ",") > 0 Then strTemp = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") - 1)
            .Rows = .Rows + 1: .MergeRow(.Rows - 1) = True
            If InStr(1, strTemp, vbTab) = 0 Then
                .RowData(.Rows - 1) = 1
            Else
                .RowData(.Rows - 1) = 2
                strTemp = Mid(strTemp, 2)
            End If
            Set .Cell(flexcpPicture, .Rows - 1, mCol.����) = Me.imgList.ListImages(.RowData(.Rows - 1)).Picture
            .TextMatrix(.Rows - 1, mCol.����) = Mid(strTemp, 2)
            .TextMatrix(.Rows - 1, mCol.���� + 1) = .TextMatrix(.Rows - 1, mCol.����)
            If Val(Left(strTemp, 1)) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.��Ӱ) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.��Ӱ) = flexUnchecked
            End If
            If InStr(1, aryItem(lngCount), ",") > 0 Then
                strTemp = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") + 1)
                aryChild = Split(strTemp, ",")
                For lngChild = 0 To UBound(aryChild)
                    strTemp = aryChild(lngChild)
                    .Rows = .Rows + 1: .MergeRow(.Rows - 1) = True
                    .RowData(.Rows - 1) = 2
                    Set .Cell(flexcpPicture, .Rows - 1, mCol.���� + 1) = Me.imgList.ListImages(.RowData(.Rows - 1)).Picture
                    .TextMatrix(.Rows - 1, mCol.���� + 1) = Mid(strTemp, 2)
                    If Val(Left(strTemp, 1)) = 1 Then
                        .Cell(flexcpChecked, .Rows - 1, mCol.��Ӱ) = flexChecked
                    Else
                        .Cell(flexcpChecked, .Rows - 1, mCol.��Ӱ) = flexUnchecked
                    End If
                Next
            End If
        Next

        If .Rows > .FixedRows Then .Row = .FixedRows
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(strKind As String, strPart As String) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    mstrKind = strKind: mstrPart = strPart
    
    '�����ǰ��Ŀ����ʾ
    Me.txt����.Text = "": Me.txt����.Text = "": Me.cbo����.Text = "": Me.txt��ע.Text = ""
    If mstrPart = "" Then FormatList: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    err = 0: On Error GoTo ErrHand
    gstrSql = "Select ����, ����, ����, ��ע, ����,�����Ա� From ���Ƽ�鲿λ Where ���� = [1] And ���� = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind, mstrPart)
    With rsTemp
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.cbo����.Tag = .Fields("����").DefinedSize
        Me.cbo����.Tag = .Fields("����").DefinedSize
        Me.txt��ע.MaxLength = .Fields("��ע").DefinedSize
        If .RecordCount > 0 Then
            Me.txt����.Text = "" & !����
            Me.txt����.Text = "" & !����
            Me.cbo����.Text = "" & !����
            Me.txt��ע.Text = "" & !��ע
            Me.cbo�����Ա�.ListIndex = NVL(!�����Ա�, 0)
            Call FormatList("" & !����)
        End If
    End With
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, strPart As String) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       strPart-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    Dim strValue As String

    err = 0: On Error GoTo ErrHand
    gstrSql = "Select Distinct ���� From ���Ƽ�鲿λ Where ���� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind)
    With rsTemp
        strValue = Me.cbo����.Text
        Me.cbo����.Clear
        Do While Not .EOF
            Me.cbo����.AddItem "" & !����
            .MoveNext
        Loop
        Me.cbo����.Text = strValue
    End With
    
    gstrSql = "Select ���� From Table(Cast(f_Check_Motheds([1]) As " & gstrDBOwner & ".t_Dic_Rowset)) Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind)
    With rsTemp
        strValue = Me.cbo����.Text
        Me.cbo����.Clear
        Do While Not .EOF
            Me.cbo����.AddItem "" & !����
            .MoveNext
        Loop
        Me.cbo����.Text = strValue
    End With
    
   
    
    If blnAdd Then
        gstrSql = "Select Nvl(Max(����), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From ���Ƽ�鲿λ Where ���� = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind)
        With rsTemp
            If !���� <> 0 And !���� <= Me.txt����.MaxLength Then
                Me.txt����.Text = Format(Val(!����) + 1, String(!����, "0"))
            Else
                Me.txt����.Text = Format(Val(!����) + 1, String(Me.txt����.MaxLength, "0"))
            End If
        End With
        
        '��������ñ�עֵ
        Me.txt����.Text = "": Me.txt��ע.Text = ""
        'Call FormatList
    End If
    
    mstrPart = strPart
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "����", "�޸�"): Call Form_Resize
    Me.txt����.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    Call zlRefresh(mstrKind, mstrPart)
End Sub

Public Function zlEditSave() As String
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀ����,����ʧ�ܷ���""
    Dim strOption As String, strCheck As String, lngCheck As Long
    Dim blnTrans As Boolean, blnRisTrans As Boolean
    Dim strUpOption As String, strUpCheck As String
        
    With Me.vfg����
        strOption = ""
        strUpOption = ""
        For lngCount = .FixedRows To .Rows - 1
            If .RowData(lngCount) = 1 Then
                strOption = strOption & ";" & IIf(.Cell(flexcpChecked, lngCount, mCol.��Ӱ) = flexChecked, 1, 0)
                strOption = strOption & .TextMatrix(lngCount, mCol.����)
                strUpOption = strUpOption & ";|"
                strUpOption = strUpOption & .TextMatrix(lngCount, mCol.����)
                strCheck = ""
                strUpCheck = ""
                For lngCheck = lngCount + 1 To .Rows - 1
                    If .TextMatrix(lngCheck, mCol.����) <> "" Then Exit For
                    strCheck = strCheck & "," & IIf(.Cell(flexcpChecked, lngCheck, mCol.��Ӱ) = flexChecked, 1, 0)
                    strCheck = strCheck & .TextMatrix(lngCheck, mCol.���� + 1)
                    strUpCheck = strUpCheck & "," & .TextMatrix(lngCount, mCol.����) & "|" & .TextMatrix(lngCheck, mCol.���� + 1)
                Next
                strOption = strOption & strCheck
                strUpOption = strUpOption & strUpCheck
            End If
            
            If .RowData(lngCount) = 2 And .TextMatrix(lngCount, mCol.����) <> "" Then
                strUpCheck = ""
                strCheck = ""
                strCheck = strCheck & vbTab & IIf(.Cell(flexcpChecked, lngCount, mCol.��Ӱ) = flexChecked, 1, 0)
                strCheck = strCheck & .TextMatrix(lngCount, mCol.����)
                strUpCheck = strUpCheck & ";|" & .TextMatrix(lngCount, mCol.����)
                
                If strCheck <> "" Then strOption = strOption & vbTab & Mid(strCheck, 2)
                If strUpCheck <> "" Then strUpOption = strUpOption & ";" & Mid(strUpCheck, 2)
                strCheck = ""
                strUpCheck = ""
                For lngCheck = lngCount + 1 To .Rows - 1
                    If .TextMatrix(lngCheck, mCol.����) <> "" Then Exit For
                    strCheck = strCheck & "," & IIf(.Cell(flexcpChecked, lngCheck, mCol.��Ӱ) = flexChecked, 1, 0)
                    strCheck = strCheck & .TextMatrix(lngCheck, mCol.���� + 1)
                    strUpCheck = strUpCheck & "," & .TextMatrix(lngCount, mCol.����) & "|" & .TextMatrix(lngCheck, mCol.���� + 1)
                Next
                strOption = strOption & strCheck
                strUpOption = strUpOption & strUpCheck
            End If
        Next
'        If strOption <> "" Then strOption = Mid(strOption, 2)
        
        If strOption = "" Then
            MsgBox "����������һ�ּ�鷽����", vbInformation, gstrSysName
            .SetFocus: zlEditSave = "": Exit Function
        End If
    End With
    
    
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    If Val(Me.txt����.Text) > Val(String(Me.txt����.MaxLength, "9")) Then
        MsgBox "����̫��", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = "": Exit Function
    End If
    If Trim(Me.cbo����.Text) = "" Then
        MsgBox "��������飡", vbInformation, gstrSysName
        Me.cbo����.SetFocus: zlEditSave = "": Exit Function
    End If
    If LenB(StrConv(Trim(Me.cbo����.Text), vbFromUnicode)) > Val(Me.cbo����.Tag) Then
        MsgBox "���鳬�������" & Val(Me.cbo����.Tag) & "���ַ�����", vbInformation, gstrSysName
        Me.cbo����.SetFocus: zlEditSave = "": Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt��ע.Text), vbFromUnicode)) > Me.txt��ע.MaxLength Then
        MsgBox "��ע���������" & Me.txt��ע.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt��ע.SetFocus: zlEditSave = "": Exit Function
    End If
    
    '���ݱ��������֯
    gstrSql = "'" & mstrKind & "','" & mstrPart & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.cbo����.Text) & "','" & Trim(Me.txt��ע.Text) & "','" & strOption & "'"
    gstrSql = gstrSql & "," & Me.cbo�����Ա�.ListIndex & ",'" & strUpOption & "'"
    
    If Me.Tag = "����" Then
        gstrSql = "Zl_���Ƽ�鲿λ_Edit(1," & gstrSql & ")"
    Else
        gstrSql = "Zl_���Ƽ�鲿λ_Edit(2," & gstrSql & ")"
    End If
    
    err = 0: On Error GoTo ErrHand
    
    If Me.Tag <> "����" Then
        '����RIS�ӿڣ���鲿λ�޸�ʱ����ɾ��ԭ��λ��Ӧ��������Ŀ��λ�����ò������ӿڲ�����Ч��ǰ����
        '�ŵ�HISִ�й���֮ǰ
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                Dim strSql As String
                Dim rsData As ADODB.Recordset
            
                strSql = "Select ���� From ���Ƽ�鲿λ Where ���� = [2] And ���� = [1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSql, "ȡԭ��λ����", mstrKind, mstrPart)
                If rsData.RecordCount > 0 Then
                    '���벿λ���ͺ�ԭ��λ����
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.Delete, mstrKind & "|" & rsData!����) <> 1 Then
                        '����ʱ��ʾ�ӿڴ�����Ϣ
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                        End If
                        
                        Exit Function
                    End If
                        
                    blnRisTrans = True
                End If
            Else
                '�ӿڲ�����Чʱ��ֹ����ʾ
                 MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                 
                 Exit Function
            End If
        End If
    End If
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    If Me.Tag <> "����" Then
        '����RIS�ӿڣ���鲿λ�޸�ʱ����ɾ����ԭ��λ��Ӧ��������Ŀ��λ��ǰ�����������²�λ��Ӧ�ķ��������ò������ӿڲ�����Ч��ǰ����
        '�ŵ�HISִ�й���֮��
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                '���벿λ���ͺ��²�λ����
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.AddNew, mstrKind & "|" & Trim(Me.txt����.Text)) <> 1 Then
                    gcnOracle.RollbackTrans
                    
                    '����ʱ��ʾ�ӿڴ�����Ϣ
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                    End If
                    
                    Exit Function
                End If
                    
                blnRisTrans = True
            Else
                gcnOracle.RollbackTrans
                
                '�ӿڲ�����Чʱ��ֹ����ʾ
                 MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                 
                 Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTrans = False
    blnRisTrans = False
    
    mstrPart = Trim(Me.txt����.Text)
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    zlEditSave = mstrPart: Exit Function
    
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    
    'Ris�ӿں�HIS��ͬ��ʱ��д������־
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HIS������ɾ����鲿λ����RIS�ӿں�HIS���ݲ�ͬ��������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmCheckPartList��cbsThis_Execute", "HIS������ɾ����鲿λ����RIS�ӿں�HIS���ݲ�ͬ��", "����=" & mstrKind & " " & "��λ����=" & Trim(Me.txt����.Text), 0)
    End If
    
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = "": Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
Private Sub cbo����_GotFocus()
    Me.cbo����.SelStart = 0: Me.cbo����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR & ",;", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cbo����_GotFocus()
    Me.cbo����.SelStart = 0: Me.cbo����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngRowAdd As Long, lngRowParent As Long
    Dim i As Long
    
    With Me.vfg����
        lngRowAdd = 0
        If Index = 3 Then
            'ɾ������
            If .TextMatrix(.Row, mCol.����) = "" Then
                .RemoveItem .Row
            Else
                .RemoveItem .Row
                Do While .TextMatrix(.Row, mCol.����) = "" And .Row <= .Rows - 1
                    .RemoveItem .Row
                Loop
            End If
            Me.cbo����.SetFocus
            Exit Sub
        End If
        
        '��Ӵ���
        Me.cbo����.Text = Replace(Me.cbo����.Text, ",", "")
        Me.cbo����.Text = Replace(Me.cbo����.Text, ";", "")
        If Trim(Me.cbo����.Text) = "" Then MsgBox "��ָ���������ƣ�", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.cbo����.Text), vbFromUnicode)) > Val(Me.cbo����.Tag) Then
            MsgBox "��鷽�����������" & Val(Me.cbo����.Tag) & "���ַ�����", vbInformation, gstrSysName
            Me.cbo����.SetFocus: Exit Sub
        End If
        Select Case Index
        Case 0
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.���� + 1) = Trim(Me.cbo����.Text) Then
                    MsgBox "�Ѿ������˸÷����������ظ���", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
                End If
'                If .TextMatrix(lngCount, mCol.����) <> "" And .RowData(lngCount) = 2 And lngRowAdd = 0 Then lngRowAdd = lngCount
            Next
            If lngRowAdd = 0 Then lngRowAdd = .Rows
            .AddItem Trim(Me.cbo����.Text) & vbTab & Trim(Me.cbo����.Text), lngRowAdd
            .Row = lngRowAdd: .RowData(.Row) = 1: .MergeRow(.Row) = True
            Set .Cell(flexcpPicture, .Row, mCol.����) = Me.imgList.ListImages(.RowData(.Row)).Picture
        Case 1
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.���� + 1) = Trim(Me.cbo����.Text) Then
                    MsgBox "�Ѿ������˸÷����������ظ���", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
                End If
            Next
            lngRowAdd = .Rows
            .AddItem Trim(Me.cbo����.Text) & vbTab & Trim(Me.cbo����.Text), .Rows
            .Row = .Rows - 1: .RowData(.Row) = 2: .MergeRow(.Row) = True
            Set .Cell(flexcpPicture, .Row, mCol.����) = Me.imgList.ListImages(.RowData(.Row)).Picture
        Case 2
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.����) = Trim(Me.cbo����.Text) Then
                    MsgBox "�Ѿ������˸÷����������ظ���", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
                End If
            Next
            lngRowParent = .Row
            If .TextMatrix(.Row, mCol.����) = "" Then
                For lngCount = .Row - 1 To .FixedRows Step -1
                    If .TextMatrix(lngCount, mCol.����) <> "" Then lngRowParent = lngCount: Exit For
                Next
            Else
                lngRowParent = .Row
            End If
            For lngCount = lngRowParent + 1 To .Rows - 1
                If .TextMatrix(lngCount, mCol.����) <> "" Then lngRowAdd = lngCount: Exit For
                If .TextMatrix(lngCount, mCol.���� + 1) = Trim(Me.cbo����.Text) Then
                    MsgBox "�Ѿ������˸÷����������ظ���", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
                End If
            Next
            '���÷����²���������ͬ�ĸ��ӷ���
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.����) <> "" And .RowData(lngCount) <> .RowData(.Row) Then
                    For i = lngCount + 1 To .Rows - 1
                        If .TextMatrix(i, mCol.����) <> "" Then Exit For
                        If .TextMatrix(i, 1) = Me.cbo����.Text Then
                            MsgBox "���ÿ�ѡ�����²����������ͬ�ĸ��ӷ�����", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
                        End If
                    Next
                End If
            Next
            If lngRowAdd = 0 Then lngRowAdd = .Rows
            .AddItem "" & vbTab & Trim(Me.cbo����.Text), lngRowAdd
            .Row = lngRowAdd: .RowData(.Row) = 2: .MergeRow(.Row) = True
            Set .Cell(flexcpPicture, .Row, mCol.���� + 1) = Me.imgList.ListImages(.RowData(.Row)).Picture
        End Select
        If InStr(1, Trim(Me.cbo����.Text), "��ǿ") > 0 Or InStr(1, Trim(Me.cbo����.Text), "��Ӱ") > 0 Then
            .Cell(flexcpChecked, .Row, mCol.��Ӱ) = flexChecked
        Else
            .Cell(flexcpChecked, .Row, mCol.��Ӱ) = flexUnchecked
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        Me.cbo����.SetFocus
    End With
    Call vfg����_RowColChange
End Sub

Private Sub Form_Load()
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    
    Me.cbo�����Ա�.Clear
    With Me.cbo�����Ա�
        .AddItem "0-���Ա�����"
        .AddItem "1-����"
        .AddItem "1-Ů��"
    End With
    
    Me.cbo�����Ա�.ListIndex = 0
    
    mstrPart = ""
    Call FormatList
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.Tag <> "" Then
        Me.BackColor = RGB(230, 230, 230)
        Me.vfg����.FocusRect = flexFocusHeavy
        Me.cmdEdit(0).Enabled = True
        Me.cmdEdit(1).Enabled = True
        Me.cmdEdit(2).Enabled = True
        Me.cmdEdit(3).Enabled = True
    Else
        Me.BackColor = &H8000000F
        Me.vfg����.FocusRect = flexFocusNone
        Me.cmdEdit(0).Enabled = False
        Me.cmdEdit(1).Enabled = False
        Me.cmdEdit(2).Enabled = False
        Me.cmdEdit(3).Enabled = False
    End If
End Sub

Private Sub txt��ע_GotFocus()
    Me.txt��ע.SelStart = 0: Me.txt��ע.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfg����_DblClick()
    If Me.vfg����.MouseRow < Me.vfg����.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    With Me.vfg����
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.��Ӱ) = flexChecked Then
            .Cell(flexcpChecked, .Row, mCol.��Ӱ) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, mCol.��Ӱ) = flexChecked
        End If
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.���� + 1) = .TextMatrix(.Row, mCol.���� + 1) Then
                .Cell(flexcpChecked, lngCount, mCol.��Ӱ) = .Cell(flexcpChecked, .Row, mCol.��Ӱ)
            End If
        Next
    End With
End Sub

Private Sub vfg����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfg����_DblClick
End Sub

Private Sub vfg����_RowColChange()
    With Me.vfg����
        If .Row < .FixedRows Then
            Me.cbo����.Text = ""
            Me.cmdEdit(0).Enabled = Me.Enabled
            Me.cmdEdit(1).Enabled = Me.Enabled
            Me.cmdEdit(2).Enabled = False
            Me.cmdEdit(3).Enabled = False
        Else
            Me.cbo����.Text = .TextMatrix(.Row, mCol.���� + 1)
            Me.cmdEdit(0).Enabled = Me.Enabled
            Me.cmdEdit(1).Enabled = Me.Enabled
            Me.cmdEdit(2).Enabled = Me.Enabled
            Me.cmdEdit(3).Enabled = Me.Enabled
        End If
    End With
End Sub

