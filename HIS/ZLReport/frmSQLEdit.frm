VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSQLEdit 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "frmSQLEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9360
   Begin XtremeSyntaxEdit.SyntaxEdit SyntaxEditSQL 
      Height          =   1485
      Left            =   240
      TabIndex        =   26
      Top             =   1320
      Width           =   5850
      _Version        =   983043
      _ExtentX        =   10319
      _ExtentY        =   2619
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   -1  'True
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
   End
   Begin VB.PictureBox picHistory 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9360
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4455
      Visible         =   0   'False
      Width           =   9360
      Begin VB.ComboBox cbocmp 
         Height          =   300
         ItemData        =   "frmSQLEdit.frx":014A
         Left            =   960
         List            =   "frmSQLEdit.frx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   305
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.CommandButton cmdCmp 
         Caption         =   "�Ա�(&D)"
         Height          =   350
         Left            =   6360
         TabIndex        =   23
         Top             =   280
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�(&E)"
         Height          =   350
         Left            =   7920
         TabIndex        =   21
         Top             =   280
         Width           =   1100
      End
      Begin VB.Frame fra2 
         Height          =   30
         Left            =   -30
         TabIndex        =   20
         Top             =   135
         Width           =   9555
      End
      Begin VB.Label lblcmp 
         Caption         =   "��ǰSQL��"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   0
      ScaleHeight     =   1245
      ScaleWidth      =   9360
      TabIndex        =   13
      Top             =   0
      Width           =   9360
      Begin VB.CommandButton cmdConn 
         Caption         =   "��"
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   930
         Width           =   300
      End
      Begin VB.ComboBox cboConn 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   930
         Width           =   1575
      End
      Begin VB.CommandButton cmdCustomProc 
         Caption         =   "��"
         Height          =   300
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   405
         Width           =   300
      End
      Begin VB.ComboBox cboData 
         Height          =   300
         ItemData        =   "frmSQLEdit.frx":01B6
         Left            =   2280
         List            =   "frmSQLEdit.frx":01B8
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   930
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cboHistory 
         Height          =   300
         ItemData        =   "frmSQLEdit.frx":01BA
         Left            =   4920
         List            =   "frmSQLEdit.frx":01C4
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label lblHistory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸ļ�¼(&H)"
         Height          =   180
         Left            =   4920
         TabIndex        =   34
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Դ(&D)"
         Height          =   180
         Left            =   2280
         TabIndex        =   33
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lblConn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&N)"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSQLEdit.frx":0226
         Height          =   615
         Left            =   1110
         TabIndex        =   14
         Top             =   75
         Width           =   8115
      End
      Begin VB.Label lblSynTest 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   960
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblCustomProc 
         AutoSize        =   -1  'True
         Caption         =   "���볣�ú���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7845
         TabIndex        =   27
         Top             =   510
         Width           =   1170
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   330
         Picture         =   "frmSQLEdit.frx":02E5
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   5265
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "����"
      FontSize        =   9
      Min             =   9
   End
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      ScaleHeight     =   1440
      ScaleWidth      =   9360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5220
      Width           =   9360
      Begin VB.TextBox txtNote 
         Height          =   720
         Left            =   4440
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   60
         Width           =   4740
      End
      Begin VB.PictureBox picCmd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   165
         ScaleHeight     =   345
         ScaleWidth      =   8895
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   990
         Width           =   8895
         Begin VB.CommandButton cmdSQLRtf 
            Caption         =   "�Ż�����"
            Height          =   350
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdPlan 
            Caption         =   "ִ�мƻ�(&P)"
            Height          =   350
            Left            =   5160
            TabIndex        =   16
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdFont 
            BackColor       =   &H00C0C0C0&
            Caption         =   "����(&F)"
            Height          =   350
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdPar 
            Caption         =   "����(&P)"
            Height          =   350
            Left            =   2700
            TabIndex        =   3
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdVerify 
            Caption         =   "��֤(&V)"
            Height          =   350
            Left            =   3960
            TabIndex        =   4
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ȷ��(&O)"
            Height          =   350
            Left            =   6540
            TabIndex        =   5
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "ȡ��(&C)"
            Height          =   350
            Left            =   7800
            TabIndex        =   6
            Top             =   0
            Width           =   1100
         End
      End
      Begin VB.Frame fra 
         Height          =   30
         Left            =   0
         TabIndex        =   11
         Top             =   855
         Width           =   9555
      End
      Begin VB.ComboBox cboType 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2850
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   690
         MaxLength       =   20
         TabIndex        =   0
         Top             =   60
         Width           =   2850
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         Height          =   180
         Left            =   3840
         TabIndex        =   15
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   225
         TabIndex        =   10
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   225
         TabIndex        =   9
         Top             =   165
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmSQLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mobjData As RPTData '�޸�:��;����/�޸�:��
Private mobjDatas As RPTDatas '�룺��ǰ��������Դ��(ֻ��)
Private mlngType As Long  '0-SQL�༭��1-SQL��ʷ���ݲ�ѯ
Private mobjPars As RPTPars '��ʱ������
Private strPreSQL As String
Private WithEvents mfrmCommProc As frmCommProc
Attribute mfrmCommProc.VB_VarHelpID = -1

Private mblnOK As Boolean

Private mlngSys As Long
Private mstrSQLCheck As String  '��ǰsql
Private mblnCheck As Boolean
Private mColSQL As Collection
Private Declare Function GetCaretPos Lib "user32" (lpPoint As PointAPI) As Long

Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByRef objData As RPTData, ByRef objDatas As RPTDatas, _
                        Optional ByVal lngType As Long)
    Set mfrmParent = frmParent
    Set mobjData = objData
    Set mobjDatas = objDatas
    mlngType = lngType
    mlngSys = lngSys
    
    Me.Show 1, frmParent
    Set objData = mobjData
    Set objDatas = mobjDatas
    ShowMe = mblnOK
End Function

Private Sub cboConn_Click()
    If Val(cboConn.Tag) = cboConn.ListIndex Then Exit Sub
    cboConn.Tag = CStr(cboConn.ListIndex)
    
    '����״̬
    If Me.Visible Then
        cmdOK.Enabled = False
        cmdPlan.Enabled = (cboConn.ListIndex = 0) And strPreSQL = SyntaxEditSQL.Text
    End If
End Sub

Private Sub cboData_Click()
    Call LoadHistory(cboData.List(cboData.ListIndex))
    Call SetComboxConnect
End Sub

Private Sub cboHistory_Click()
    SyntaxEditSQL.Text = mColSQL(cboHistory.List(cboHistory.ListIndex))
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdCmp_Click()
    Dim strThisSQL As String, strNewSQL As String
    
    If cbocmp.List(cbocmp.ListIndex) = "��ǰ���µ�����ԴSQL" Then
        strNewSQL = mobjData.SQL
    Else
        strNewSQL = mColSQL(cbocmp.List(cbocmp.ListIndex))
    End If
    strThisSQL = mColSQL(cboHistory.List(cboHistory.ListIndex))
    
    ShowDiff strThisSQL, strNewSQL
End Sub

Private Sub cmdConn_Click()
    Dim blnModified As Boolean
    Dim intIndex As String
    
    If Me.Visible = False Then Exit Sub
    
    If gfrmDBConnect Is Nothing Then
        MsgBox "�����������ӹ���ʧ�ܣ�", vbInformation, App.Title
        Exit Sub
    End If
    
    intIndex = cboConn.ListIndex
    If gfrmDBConnect.ShowMe(Me, blnModified) Then
        If blnModified Then
            '�������Ӽ�¼������
            Call mdlPublic.SetControlDBConnect(grsConnect)
            '���µ�ǰ����
            cboConn.Clear
            cboConn.AddItem "��ǰ��¼"
            Call mdlPublic.SetControlDBConnect(cboConn)
            If intIndex > cboConn.ListCount Then
                cboConn.ListIndex = 0
            Else
                cboConn.ListIndex = intIndex
            End If
            '��ն��󼯺�
            Call gclsCNs.Clear
        End If
    End If
End Sub

Private Sub cmdCustomProc_Click()
    OutCustomProc
    SyntaxEditSQL.SetFocus
End Sub

Private Sub cmdExit_Click()
    cmdCancel_Click
End Sub

Private Sub cmdFont_Click()
    On Error Resume Next
    cdg.Flags = &H3 Or &H100 Or &H400 Or &H200 Or &H10000 Or &H2000
    cdg.Min = 5: cdg.Max = 72
    cdg.FontName = SyntaxEditSQL.Font.name
    cdg.FontSize = SyntaxEditSQL.Font.Size
    cdg.FontItalic = SyntaxEditSQL.Font.Italic
    cdg.FontBold = SyntaxEditSQL.Font.Bold
    cdg.FontUnderline = SyntaxEditSQL.Font.Underline
    cdg.FontStrikethru = SyntaxEditSQL.Font.Strikethrough
    cdg.CancelError = True
    cdg.ShowFont
    If Err.Number = 0 Then
        SyntaxEditSQL.Font.name = cdg.FontName
        SyntaxEditSQL.Font.Size = cdg.FontSize
        SyntaxEditSQL.Font.Italic = cdg.FontItalic
        SyntaxEditSQL.Font.Bold = cdg.FontBold
        SyntaxEditSQL.Font.Underline = cdg.FontUnderline
        SyntaxEditSQL.Font.Strikethrough = cdg.FontStrikethru
        
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontName", SyntaxEditSQL.Font.name
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontSize", SyntaxEditSQL.Font.Size
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontUnderline", SyntaxEditSQL.Font.Underline
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontItalic", SyntaxEditSQL.Font.Italic
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontBold", SyntaxEditSQL.Font.Bold
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontStrikethru", SyntaxEditSQL.Font.Strikethrough
    Else
        Err.Clear
    End If
    SyntaxEditSQL.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, intCount As Integer
    Dim objNode As Object, arrFields() As String
    Dim strCaption As String, strSource As String
    Dim blnSuccess As Boolean
    
    If txtName.Text = "" Then
        MsgBox "�����������Դ�����ƣ�", vbInformation, App.Title
        txtName.SetFocus: Exit Sub
    End If
    If TLen(txtName.Text) > txtName.MaxLength Then
        MsgBox "����Դ�����Ƴ��Ȳ��ܳ���" & txtName.MaxLength & "���ַ���" & txtName.MaxLength \ 2 & "�����֣�", vbInformation, App.Title
        txtName.SetFocus: Exit Sub
    End If
    If TLen(txtNote.Text) > txtNote.MaxLength Then
        MsgBox "����Դ˵���ĳ��Ȳ��ܳ���" & txtNote.MaxLength & "���ַ���" & txtNote.MaxLength \ 2 & "�����֣�", vbInformation, App.Title
        txtNote.SetFocus: Exit Sub
    End If
    
    '���Ʋ����ظ�
    For Each objNode In mfrmParent.tvwSQL.Nodes
        If objNode.Key <> "Root" Then
            If objNode.Parent.Key = "Root" Then
                If mdlPublic.GetStdNodeText(objNode.Text) = txtName.Text And objNode.Key <> "_" & mobjData.Key Then
                    MsgBox "������Դ��������������Դ�����ظ���", vbInformation, App.Title
                    txtName.SetFocus: Exit Sub
                End If
            End If
        End If
    Next
    
    '���ܼ��
    If mstrSQLCheck = "" Then cmdVerify_Click
    If mstrSQLCheck <> "" Then
        If CheckSQLPlan(mstrSQLCheck, , cboConn.ItemData(cboConn.ListIndex), blnSuccess) = True Then
            If MsgBox("��ǰ����Դ�п��ܴ����������⣬�Ƿ�鿴ִ�мƻ���" & vbCrLf & "�����������档", vbQuestion + vbYesNo + vbDefaultButton2, "���ܼ��") = vbYes Then
                If InStr(mfrmParent.Caption, "]") > 0 And InStrRev(mfrmParent.Caption, "��") > 0 Then
                   strCaption = Mid(mfrmParent.Caption, InStr(mfrmParent.Caption, "]") + 1)
                   If InStrB(strCaption, "��") > 0 Then
                        strCaption = Mid(strCaption, 1, InStrRev(strCaption, "��") - 1)
                   End If
                End If
                Call frmSQLPlanEx.ShowMe(Me, cboConn.ItemData(cboConn.ListIndex), mstrSQLCheck, , strCaption & "_" & txtName.Text)
                Exit Sub
            End If
        End If
        If blnSuccess = False Then
            Exit Sub
        End If
    End If
    
    mobjData.Key = txtName.Text
    mobjData.���� = txtName.Text
    mobjData.�������ӱ�� = cboConn.ItemData(cboConn.ListIndex)
    mobjData.���� = cboType.ListIndex
    mobjData.SQL = SyntaxEditSQL.Text
    mobjData.�ֶ� = SyntaxEditSQL.Tag
    mobjData.���� = txtName.Tag
    mobjData.˵�� = txtNote.Text
    
    'ֻȡʵ����Ŀ�Ĳ�������(�����ȥ��)
    Set mobjData.Pars = New RPTPars
    intCount = GetParCount(Replace(RemoveNote(SyntaxEditSQL.Text), "[ϵͳ]", mlngSys))
    For i = 1 To intCount
        With mobjPars("_" & i - 1)
            '��ʽ����
            mobjData.Pars.Add .����, .���, .����, .����, .ȱʡֵ, .��ʽ, .ֵ�б�, .����SQL, .��ϸSQL, .�����ֶ�, .��ϸ�ֶ� _
                , .����, "_" & .Key, , .�Ƿ�����
        End With
    Next
    
    '����ˢ��
    With mfrmParent.tvwSQL
        If Caption = "��������Դ" Then
            strSource = mobjData.����
            If mobjData.�������ӱ�� > 0 Then
                '��������������ʾ���ӵ�����
                For i = 1 To cboConn.ListCount
                    If cboConn.ItemData(i) = mobjData.�������ӱ�� Then
                        strSource = mobjData.���� & "��" & Split(cboConn.List(i), "��")(1) & "��"
                        Exit For
                    End If
                Next
            End If
            If mobjData.���� = 0 Then
                Set objNode = .Nodes.Add("Root", 4, "_" & mobjData.Key, strSource, "SQL_Custom")
            Else
                Set objNode = .Nodes.Add("Root", 4, "_" & mobjData.Key, strSource, "SQL_Group")
            End If
            objNode.Expanded = True
            objNode.EnsureVisible
            
            '�����ֶ�����
            If mobjData.�ֶ� <> "" Then
                arrFields = Split(mobjData.�ֶ�, "|")
                For i = 0 To UBound(arrFields)
                    Select Case Split(arrFields(i), ",")(1)
                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR '�ı���(Varchar2,Long)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "String")
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, _
                            adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, _
                            adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt '������(Numeric(a,b),Sum)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Number")
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate '������(Date)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Date")
                        Case adBinary, adVarBinary, adLongVarBinary '������(Long Raw)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Bin")
                        Case Else '����
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Other")
                    End Select
                    objNode.Tag = Split(arrFields(i), ",")(1)
                Next
            End If
        Else
            If .SelectedItem.Children = 0 Then
                Set objNode = .SelectedItem.Parent
            Else
                Set objNode = .SelectedItem
            End If
            
            If mobjData.���� = 0 Then
                objNode.Image = "SQL_Custom"
            Else
                objNode.Image = "SQL_Group"
            End If
            
            If mobjData.�������ӱ�� > 0 Then
                '��������������ʾ���ӵ�����
                strSource = mobjData.����
                For i = 1 To cboConn.ListCount
                    If cboConn.ItemData(i) = mobjData.�������ӱ�� Then
                        strSource = mobjData.���� & "��" & Split(cboConn.List(i), "��")(1) & "��"
                        Exit For
                    End If
                Next
                objNode.Text = strSource
            Else
                objNode.Text = mobjData.����
            End If
            objNode.Key = "_" & mobjData.����
            objNode.Checked = False
            
            'ɾ���ӽ��
            Do While Not objNode.Child Is Nothing
                .Nodes.Remove objNode.Child.Index
            Loop
            
            '�����½��
            '�����ֶ�����
            If mobjData.�ֶ� <> "" Then
                arrFields = Split(mobjData.�ֶ�, "|")
                For i = 0 To UBound(arrFields)
                    Select Case Split(arrFields(i), ",")(1)
                        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR '�ı���(Varchar2,Long)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "String")
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, _
                            adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, _
                            adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt '������(Numeric(a,b),Sum)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Number")
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate '������(Date)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Date")
                        Case adBinary, adVarBinary, adLongVarBinary '������(Long Raw)
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Bin")
                        Case Else '����
                            Set objNode = .Nodes.Add("_" & mobjData.Key, 4, , Split(arrFields(i), ",")(0), "Other")
                    End Select
                    objNode.Tag = Split(arrFields(i), ",")(1)
                Next
            End If
        End If
    End With
    
    mblnOK = True
    Hide
End Sub

Private Sub cmdPar_Click()
    Dim strSQL As String, strMsg As String
    
    If TrimChar(SyntaxEditSQL.Text) = "" Then
        MsgBox "��������SQL��䣡", vbInformation, App.Title
        SyntaxEditSQL.SetFocus: Exit Sub
    End If
    
    strSQL = TrimChar(RemoveNote(SyntaxEditSQL.Text))
    strSQL = Replace(strSQL, "[ϵͳ]", mlngSys)
    
    If Not CheckPars(strSQL, strMsg, mobjPars) Then
        SyntaxEditSQL.SetFocus: Exit Sub
    Else
        If strMsg <> "" Then SyntaxEditSQL.Text = strSQL
    End If
    
    If GetParCount(strSQL) = 0 Then
        MsgBox "��SQL�����û�ж��������", vbInformation, App.Title
        SyntaxEditSQL.SetFocus: Exit Sub
    End If
    
    If mobjData.�������ӱ�� <> cboConn.ItemData(cboConn.ListIndex) Then
        mobjData.�������ӱ�� = cboConn.ItemData(cboConn.ListIndex)
    End If
    Call frmParEdit.ShowMe(Me, mlngSys, mobjData, mobjDatas, mobjPars, strSQL, mfrmParent.lngRPTID)

End Sub

Private Sub cmdPlan_Click()
    Dim strCaption As String
    
    mblnCheck = False
    If mstrSQLCheck = "" Then cmdVerify_Click
    If mstrSQLCheck <> "" And mblnCheck = False Then
        If InStr(mfrmParent.Caption, "]") > 0 And InStrRev(mfrmParent.Caption, "��") > 0 Then
           strCaption = Mid(mfrmParent.Caption, InStr(mfrmParent.Caption, "]") + 1)
           If InStrB(strCaption, "��") > 0 Then
                strCaption = Mid(strCaption, 1, InStrRev(strCaption, "��") - 1)
           End If
        End If
        frmSQLPlanEx.ShowMe Me, cboConn.ItemData(cboConn.ListIndex), mstrSQLCheck, 0, strCaption & "_" & txtName.Text
    Else
        cmdPlan.Enabled = False
    End If
End Sub

Private Sub cmdSQLRtf_Click()
    Call frmSQLPlanEx.ShowMe(Me, 0, "", 1)
End Sub

Public Sub TextFindKey(ByVal strKey As String, ByVal txtCode As SyntaxEdit, ByRef lngRow As Long, ByRef lngCol As Long)
'���ܣ������ı����ҹ���
'������
'  strKey�������ַ���
'  txtCode�����ҵ�SytaxEdit����
'  lngRow(ʵ��)�����ҵ���ʼ��
'  lngCol(ʵ��)�����ҵ���ʼ��

    Dim i As Long
    Dim lngTmp As Long, lngStart As Long, lngEnd As Long
    Dim strLine As String, strString As String
    Dim blnFind As Boolean
    
    If strKey = "" Then Exit Sub

    With txtCode
        If .RowsCount <= 0 Then
            lngRow = 0
            Exit Sub
        End If
        
        strString = LCase(strKey)
        If lngRow <= 0 Then lngRow = 1
        If lngCol <= 0 Then lngCol = 1
        '����
        For i = lngRow To .RowsCount
            strLine = LCase(.RowText(i))
            If InStr(Mid(strLine, lngCol), strString) > 0 Then
                '�ҵ�ƥ���ִ�
                lngStart = InStr(Mid(strLine, lngCol), strString) + lngCol - 1
                lngEnd = lngStart + Len(strString)
                
                .CurrPos.Row = i
                .CurrPos.Col = lngStart
                
                .Selection.Start.Col = lngStart
                .Selection.End.Col = lngEnd
                .Selection.Start.Row = i
                .Selection.End.Row = i
                .ShowSelectionMargin = True
                
                '��¼��ǰ����
                lngCol = lngEnd
                lngRow = i
                blnFind = True
                
                Exit For
            Else
                '����һ�У������׿�ʼ
                lngCol = 1
            End If
        Next
        
    End With
    
End Sub

Private Sub cmdVerify_Click()
    Dim strFields As String, strObject As String
    Dim strSQL As String, strR As String
    Dim strFieldInfo As String, strMsg As String
    
    strSQL = RemoveNote(SyntaxEditSQL.Text)
    strSQL = TrimChar(strSQL)
    strSQL = Replace(strSQL, "[ϵͳ]", mlngSys)
    mblnCheck = True
    
    If strSQL = "" Then
        MsgBox "��������SQL��䣡", vbInformation, App.Title
        SyntaxEditSQL.SetFocus: Exit Sub
    End If
    
    If Not CheckPars(strSQL, strMsg, mobjPars) Then
        SyntaxEditSQL.SetFocus: Exit Sub
    Else
        If strMsg <> "" Then SyntaxEditSQL.Text = strSQL
    End If
    
    If GetParCount(strSQL) > mobjPars.count Then '���ӵĲ�����������,����Ĳ���ȷ��ʱ�Զ�ɾ��
        MsgBox "SQL����д��ڶ����˵���δ���õĲ���,�������ò�����", vbInformation, App.Title
        cmdPar.SetFocus: Exit Sub
    End If
    
    'SQL����������Ȩ�޼��(DBLINK����@�ģ���������Ȩ��)
    'ȡ����
    strObject = SQLObject(strSQL)
    If strObject = "" And InStr(UCase(strSQL), "TABLE") = 0 And InStr(UCase(strSQL), "@") = 0 Then
        MsgBox "���ܷ���SQL�������ѯ�����ݶ���,�����Ƿ���ȷ��д��", vbInformation, App.Title
        SyntaxEditSQL.SetFocus: Exit Sub
    End If
    
    '�Ƿ���Ȩ��
    If cboConn.ItemData(cboConn.ListIndex) = Val("0-��ǰ��¼����") Then
        strR = CheckObjectPriv(strObject)
        If strR <> "" Then
            MsgBox "���ж��󲻴��ڻ�û��Ȩ�޷���:" & vbCrLf & strR, vbInformation, App.Title
            SyntaxEditSQL.SetFocus
            Call TextFindKey(Split(strR, ",")(0), SyntaxEditSQL, 1, 1)
            SyntaxEditSQL.RefreshColors
            Exit Sub
        End If
    Else
        cmdPlan.Enabled = False
    End If
    
    'ȡ������
    strObject = ObjectOwner(strObject, Me, cboConn.ItemData(cboConn.ListIndex))
    If strObject = "ȡ��" Then Exit Sub 'ȡ������
    
    strSQL = SQLOwner(strSQL, strObject)
    
    ShowFlash "����У������Դ��ȷ��,���Ժ� ..."
    
    If GetParCount(strSQL) = 0 Then
        strFields = CheckSQL(strSQL, strR, , mstrSQLCheck, strFieldInfo, mobjDatas _
                            , cboConn.ItemData(cboConn.ListIndex))
    Else
        strFields = CheckSQL(strSQL, strR, ReplaceParSysNo(mobjPars, mlngSys), mstrSQLCheck, strFieldInfo, mobjDatas _
                            , cboConn.ItemData(cboConn.ListIndex))
    End If
    
    ShowFlash
    
    If strFields = "" Then
        MsgBox "SQL���У��ʧ�ܣ�" & vbCrLf & vbCrLf & _
            "���� " & strR & vbCrLf & vbCrLf & _
            "�����Ƿ���ȷ��д,������Ƿ���ȷ���ã�", vbInformation, App.Title
        SyntaxEditSQL.SetFocus
        If InStr(UCase(SyntaxEditSQL.Text), strFieldInfo) > 0 Then
            SyntaxEditSQL.CurrPos.StrPos = InStr(UCase(SyntaxEditSQL.Text), strFieldInfo) - 1
            SyntaxEditSQL.Selection.Start.StrPos = SyntaxEditSQL.CurrPos.StrPos
            SyntaxEditSQL.Selection.End.StrPos = SyntaxEditSQL.Selection.Start.StrPos + Len(strFieldInfo)
            SyntaxEditSQL.RefreshColors
        End If
    Else
        SyntaxEditSQL.Tag = strFields
        txtName.Tag = strObject
        strPreSQL = SyntaxEditSQL.Text
                
        'ȱʡ����Դ��
        If txtName.Text = "" Then
            If strObject <> "" Then txtName.Text = Split(Split(strObject, ",")(0), ".")(1) & "_����"
        End If
        'ȱʡ����Դ����
        If mobjData.SQL = "" Then
            If UCase(SyntaxEditSQL.Text) Like UCase("*Group by*") Then
                cboType.ListIndex = 1
            End If
        End If
        
        cmdOK.Enabled = True
        cmdPlan.Enabled = (cboConn.ListIndex = 0)
        mblnCheck = False
    End If
End Sub

Private Sub Form_Activate()
    cmdConn.Refresh         '���ֵ�����ʾ���쳣���ʼӴ���
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub SetSyntaxEditColor(ByRef objSyntaxEdit As SyntaxEdit)
'���ܣ�����SyntaxEdit�ؼ����ı���ɫ
    
    If objSyntaxEdit.SyntaxScheme = "" Then
        With objSyntaxEdit
            .SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
            .SyntaxScheme = ""
        End With
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim objFSO As New Scripting.FileSystemObject
    
    RestoreWinState Me, App.ProductName

    SyntaxEditSQL.Font.name = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontName", "Fixedsys")
    SyntaxEditSQL.Font.Size = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontSize", 12)
    SyntaxEditSQL.Font.Underline = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontUnderline", 0)
    SyntaxEditSQL.Font.Italic = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontItalic", 0)
    SyntaxEditSQL.Font.Bold = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontBold", 0)
    SyntaxEditSQL.Font.Strikethrough = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontStrikethru", 0)
    SyntaxEditSQL.BorderStyle = xtpBorderClientEdge
    '���ÿؼ�����ʾ��ɫ����Ϊ��SQL
    If objFSO.FileExists(App.Path & "\_sql.schclass") Then
        gstrColor = ReadFileToString(App.Path & "\_sql.schclass")
    Else
        gstrColor = ""
    End If
    SyntaxEditSQL.SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
    SyntaxEditSQL.SyntaxScheme = gstrColor

    mstrSQLCheck = ""
    
    cboType.AddItem "ȱʡΪ�����ͷ����ṩ����" 'ȱʡ
    cboType.AddItem "ȱʡΪ������ܱ���ṩ����"
    
    cboConn.Clear
    cboConn.AddItem "��ǰ��¼"
    
    '��ȡ��������
    Call mdlPublic.SetControlDBConnect(cboConn)
    
    If mlngType = 0 Then
        If cboConn.ListCount <= 0 Then
            cboConn.Enabled = False
            cboConn.BackColor = Me.BackColor
        End If
        
        If mobjData Is Nothing Then
            Caption = "��������Դ"
            cboType.ListIndex = 0
            cboConn.ListIndex = 0
            cmdOK.Enabled = False
            cmdPlan.Enabled = False
            
            strPreSQL = ""
            
            Set mobjData = Nothing: Set mobjData = New RPTData
            Set mobjPars = New RPTPars
        Else
            Caption = "�޸�����Դ"
            strPreSQL = mobjData.SQL
            SyntaxEditSQL.Text = mobjData.SQL
            
            SyntaxEditSQL.Tag = mobjData.�ֶ�
            txtName.Tag = mobjData.����
            
            txtName.Text = mobjData.����
            cboType.ListIndex = mobjData.����
            
            txtNote.Text = mobjData.˵��
            
            'ͬ����ǰѡ��
            For i = 0 To cboConn.ListCount - 1
                If cboConn.ItemData(i) > 0 And cboConn.ItemData(i) = mobjData.�������ӱ�� Then
                    cboConn.ListIndex = i
                    Exit For
                End If
            Next
            If cboConn.ListIndex < 0 Then cboConn.ListIndex = 0
            
            CopyPars mobjData.Pars, mobjPars
        End If
        lblData.Visible = False
        lblHistory.Visible = False
    ElseIf mlngType = 1 Then
        Caption = "�鿴��ʷ����Դ"
        lblCaption = vbCrLf & "����������鿴����Դ��ʷ�޸ļ�¼������ԱȲ鿴��ʷSQL�����������F3�����������е�SQL��䡣"
        picHistory.Visible = True
        picDown.Visible = False
        cboHistory.Visible = True
        cbocmp.Visible = True
        cmdCmp.Visible = True
        cboData.Visible = True
        SyntaxEditSQL.ReadOnly = True
        lblcmp.Visible = True
        lblCustomProc.Visible = False
        cmdCustomProc.Visible = False
        cmdConn.Visible = False
        cboConn.Enabled = False
        cboConn.BackColor = Me.BackColor
        For i = 1 To mobjDatas.count
            cboData.AddItem mobjDatas(i).����
            If mobjDatas(i).���� = IIF(mobjData.ԭ���� = "", mobjData.����, mobjData.ԭ����) Then
                'Ĭ��ѡ������ѡ�е�����Դ
                cboData.ListIndex = cboData.NewIndex
                '������������
                Call SetComboxConnect
            End If
        Next
        If cboData.ListIndex < 0 And cboData.ListCount > 0 Then cboData.ListIndex = 0
    End If
    SyntaxEditSQL.RefreshColors
    cmdPlan.Enabled = (cboConn.ListIndex = 0)
End Sub

Private Sub LoadHistory(ByVal str����Դ���� As String)
'���ܣ���ȡ��ʷ����Դ��¼
    Dim rsTmp As Recordset, strSQL As String
    Dim strKey As String, strHisSQL As String
    
    On Error GoTo errH
    strSQL = "Select RPAD(�޸���,25) as �޸���,To_Char(�޸�ʱ��, 'yyyy-mm-dd hh24:mi:ss') �޸�ʱ��,�к�,���� " & _
             "From zlRPTSQLsHistory " & _
             "Where ����ID=[1] and ����Դ����=[2] " & _
             "Order By �޸�ʱ�� Desc,�к�"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, mfrmParent.lngRPTID, str����Դ����)
    If rsTmp.RecordCount > 0 Then
        Set mColSQL = New Collection
        cboHistory.Clear
        cbocmp.Clear
        cbocmp.AddItem "��ǰ���µ�����ԴSQL"
        Do While Not rsTmp.EOF
            If strKey = rsTmp!�޸��� & "" & rsTmp!�޸�ʱ�� Then
                strHisSQL = strHisSQL & vbCrLf & Nvl(rsTmp!����)
            Else
                If strHisSQL <> "" Then
                    mColSQL.Add strHisSQL, strKey
                    cboHistory.AddItem strKey
                    cbocmp.AddItem strKey
                End If
                strHisSQL = Nvl(rsTmp!����)
            End If
            strKey = rsTmp!�޸��� & "" & rsTmp!�޸�ʱ��
            rsTmp.MoveNext
        Loop
        If strHisSQL <> "" Then
            mColSQL.Add strHisSQL, strKey
            cboHistory.AddItem strKey
            cbocmp.AddItem strKey
        End If
    End If
    If cboHistory.ListCount > 0 Then
        cboHistory.ListIndex = 0
    End If
    cbocmp.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then mblnOK = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width <= 9600 Then Me.Width = 9600
    If Me.Height <= 7200 Then Me.Height = 7200

    picTop.Width = Me.ScaleWidth

    If mlngType = Val("1-�鿴����Դ") Then
        cboHistory.Left = Me.ScaleWidth - cboHistory.Width - 60
        cboHistory.Top = picTop.ScaleHeight - cboHistory.Height - 60
                         
        cboData.Width = (Me.ScaleWidth - cboHistory.Width) \ 2 - 60 * 3
        cboData.Left = cboHistory.Left - cboData.Width - 60 * 2
        cboData.Top = cboHistory.Top
                        
        cboConn.Left = 60
        cboConn.Top = cboHistory.Top
        cboConn.Width = cboData.Width
        
        lblConn.Left = cboConn.Left
        lblData.Left = cboData.Left
        lblHistory.Left = cboHistory.Left
    Else
        picTop.Height = 1100
        lblConn.Left = 60
        lblConn.Top = picTop.ScaleHeight - lblConn.Height - 120
        cboConn.Left = lblConn.Left + lblConn.Width + 60
        cboConn.Top = lblConn.Top - 60
        cboConn.Width = Me.ScaleWidth \ 3 - lblConn.Width - 60 - cmdConn.Width
        cmdConn.Top = cboConn.Top
        cmdConn.Left = cboConn.Left + cboConn.Width
    End If
    
    SyntaxEditSQL.Left = ScaleLeft + 30
    SyntaxEditSQL.Top = ScaleTop + picTop.Height
    SyntaxEditSQL.Width = ScaleWidth - 60
    SyntaxEditSQL.Height = ScaleHeight - IIF(picDown.Visible, picDown.Height, 0) _
                            - IIF(picHistory.Visible, picHistory.Height, 0) - picTop.Height
    
    txtNote.Width = Me.ScaleWidth - txtNote.Left - 100
    
    picDown.Width = Me.ScaleWidth
    picHistory.Width = Me.ScaleWidth
    picCmd.Width = Me.ScaleWidth
    
    fra.Left = 60
    fra.Width = Me.ScaleWidth - 60 * 2
    fra2.Left = 60
    fra2.Width = Me.ScaleWidth - 60 * 2
    
    lblCaption.Width = picTop.ScaleWidth - lblCaption.Left - 100
    
    cmdCustomProc.Left = Me.ScaleWidth - cmdCustomProc.Width - 60
    cmdCustomProc.Top = cboConn.Top
    lblCustomProc.Left = cmdCustomProc.Left - lblCustomProc.Width - 15
    lblCustomProc.Top = cboConn.Top + 60
    
    If mlngType = Val("1-�鿴����Դ") Then
        cmdExit.Left = picCmd.ScaleWidth - cmdExit.Width - 150
        cmdCmp.Left = cmdExit.Left - cmdCmp.Width - 150
    Else
        cmdCancel.Left = picCmd.ScaleWidth - picCmd.Left - cmdCancel.Width - 150
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 150
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjData = Nothing
    If Not mfrmCommProc Is Nothing Then
        Unload mfrmCommProc
    End If
    SaveWinState Me, App.ProductName
End Sub

Private Sub mfrmCommProc_AfterSelect(ByVal strProc As String)
    SyntaxEditSQL.Selection.Text = strProc
End Sub

Private Sub SyntaxEditSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTmp As String

    If KeyCode = vbKeyA And Shift = 2 Then
        SyntaxEditSQL.SelectAll 'Ctrl+A
    ElseIf KeyCode = vbKeyF3 Then
        If Len(SyntaxEditSQL.Selection.Text) = 0 Then
            strTmp = SyntaxEditSQL.Text
        Else
            strTmp = SyntaxEditSQL.Selection.Text
        End If
        strTmp = Replace(strTmp, "[ϵͳ]", mlngSys)
        
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText GetEditSQL(strTmp, ReplaceParSysNo(mobjPars, mlngSys))
    ElseIf KeyCode = vbKeyF4 Then
        strTmp = Clipboard.GetText
        
        If Replace(strTmp, "*", "@") Like "*/@B*@/*" Or Replace(strTmp, "*", "@") Like "*/@E*@/*" Then
            If Len(SyntaxEditSQL.Selection.Text) = 0 Then
                SyntaxEditSQL.Text = GetParSQL(strTmp)
            Else
                SyntaxEditSQL.Selection.Text = GetParSQL(strTmp)
            End If
        Else
            SyntaxEditSQL.Selection.Text = strTmp
        End If
    ElseIf KeyCode = vbKeyZ And Shift = 2 Then
        SyntaxEditSQL.Undo
    ElseIf KeyCode = vbKeyY And Shift = 2 Then
        SyntaxEditSQL.Redo
    ElseIf KeyCode = vbKeyC And Shift = 2 Then
        SyntaxEditSQL.Copy
    ElseIf KeyCode = vbKeyV And Shift = 2 Then
        SyntaxEditSQL.Paste
    ElseIf KeyCode = vbKeyF6 Then
        OutCustomProc
    End If
End Sub

Private Sub OutCustomProc()
    Dim lngSelstart As Long
    Dim lngPointTop As Long
    Dim lngPointLeft As Long
    Dim lngParentLeft As Long
    Dim lngParentTop As Long
    Dim objPoint As PointAPI
    Dim rParent As RECT
    lblSynTest.FontName = SyntaxEditSQL.Font.name
    lblSynTest.FontSize = SyntaxEditSQL.Font.Size
    lblSynTest.FontItalic = SyntaxEditSQL.Font.Italic
    lblSynTest.FontBold = SyntaxEditSQL.Font.Bold
    lblSynTest.FontUnderline = SyntaxEditSQL.Font.Underline
    lblSynTest.FontStrikethru = SyntaxEditSQL.Font.Strikethrough
    
    '�ȱ����¹��������
    lngSelstart = SyntaxEditSQL.CurrPos.Col
    '������֮ǰ���ı�����
    Call SyntaxEditSQL.Selection.Start.SetPos(SyntaxEditSQL.CurrPos.Row, 0)
    Call SyntaxEditSQL.Selection.End.SetPos(SyntaxEditSQL.CurrPos.Row, lngSelstart)
    '��ֵ�ı�����
    lblSynTest.Caption = SyntaxEditSQL.Selection.Text
    '����������λ�õ�����
    Call GetCaretPos(objPoint)
    lngPointLeft = objPoint.X * Screen.TwipsPerPixelX
    lngPointTop = objPoint.Y * Screen.TwipsPerPixelY + lblSynTest.Height
    '���㴰�������Ļ������
    GetWindowRect SyntaxEditSQL.hwnd, rParent
    lngParentLeft = rParent.Left * Screen.TwipsPerPixelX
    lngParentTop = rParent.Top * Screen.TwipsPerPixelY
    '���¶�λ���
    Call SyntaxEditSQL.Selection.Start.SetPos(SyntaxEditSQL.CurrPos.Row, lngSelstart)
    If mfrmCommProc Is Nothing Then
        Set mfrmCommProc = New frmCommProc
    End If
    '�ж��Ƿ񵯳���ĸ߶ȳ���������ײ�
    If SyntaxEditSQL.Height + picDown.Height - lngPointTop < mfrmCommProc.Height Then
        lngPointTop = lngPointTop - mfrmCommProc.Height
    End If
    Call mfrmCommProc.ShowMe(Me, lngParentLeft + lngPointLeft, lngParentTop + lngPointTop)
End Sub

Private Sub SyntaxEditSQL_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    If SyntaxEditSQL.Text <> strPreSQL Then
        cmdOK.Enabled = False
        cmdPlan.Enabled = False
    ElseIf TrimChar(SyntaxEditSQL.Text) <> "" Then
        cmdOK.Enabled = True
        cmdPlan.Enabled = True
    End If
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr("|.'`~!@#$^&{}"";:\����" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 And txtName.Text <> "" Then
        KeyAscii = 0: SendKeys "{Tab}"
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub txtNote_GotFocus()
    SelAll txtNote
End Sub

Private Sub txtName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtName.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtName.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtName.hwnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub SetComboxConnect()
    If mlngType = Val("1-�鿴����Դ") Then
        '������������ѡ��
        Dim i As Integer, j As Integer
        
        For i = 1 To mobjDatas.count
            If UCase(Trim(mobjDatas.Item(i).����)) = UCase(Trim(cboData.List(cboData.ListIndex))) Then
                For j = 0 To cboConn.ListCount - 1
                    If mobjDatas.Item(i).�������ӱ�� = cboConn.ItemData(j) Then
                        cboConn.ListIndex = j
                        Exit Sub
                    End If
                Next
            End If
        Next
    End If
End Sub
