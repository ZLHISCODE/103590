VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceCopy 
   AutoRedraw      =   -1  'True
   Caption         =   "����ҽ��"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmAdviceCopy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   465
      TabIndex        =   8
      ToolTipText     =   "F1"
      Top             =   5730
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   5010
      Left            =   15
      TabIndex        =   0
      Top             =   555
      Width           =   9435
      _cx             =   16642
      _cy             =   8837
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceCopy.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
      Editable        =   2
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   4
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6090
      TabIndex        =   3
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   2745
      TabIndex        =   2
      ToolTipText     =   "Ctrl+R"
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   1635
      TabIndex        =   1
      ToolTipText     =   "Ctrl+A"
      Top             =   5730
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   615
      Left            =   15
      TabIndex        =   9
      Top             =   -75
      Width           =   9420
      Begin VB.CommandButton cmdPati 
         Height          =   240
         Left            =   1740
         Picture         =   "frmAdviceCopy.frx":077D
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����(F4)"
         Top             =   225
         Width           =   255
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   195
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&P)"
         Height          =   180
         Left            =   135
         TabIndex        =   11
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "###������Ϣ###"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2130
         TabIndex        =   10
         Top             =   255
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3975
      Left            =   795
      TabIndex        =   7
      Top             =   450
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "סԺ��"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "סԺҽʦ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�Ա�"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "�ѱ�"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����ȼ�"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   $"frmAdviceCopy.frx":0873
         Object.Width           =   2081
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "��Ժ����"
         Object.Width           =   2081
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "���ʽ"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2295
      Top             =   210
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
            Picture         =   "frmAdviceCopy.frx":0880
            Key             =   "Pati"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdviceCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object
Private mstrPrivs As String
Private mbln��ʿվ As Boolean
Private mlngǰ��ID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstr�Һŵ� As String
Private mblnMoved As Boolean
Private mblnItem As Boolean
Private mstrIDs As String
Private mstrAlter As String

Private Enum COL���׷���
    colѡ�� = 0
    col��Ч = 1
    colʱ�� = 2
    col���� = 3
    col���� = 4
    col������λ = 5
    col���� = 6
    col������λ = 7
    colƵ�� = 8
    col�÷� = 9
    col���� = 10
    colִ��ʱ�� = 11
    colִ�п��� = 12
    colID = 13
    col���ID = 14
    col������� = 15
    col������ĿID = 16
    col�շ�ϸĿID = 17
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, _
    lng����ID As Long, varTime As Variant, blnMoved As Boolean, _
    Optional ByVal bln��ʿվ As Boolean, Optional ByVal lngǰ��ID As Long, _
    Optional strAlter As String) As String
'���أ�lng����ID,varTime=Ҫ����ҽ���Ĳ���ID����ҳID(�Һŵ�NO)
'      blnMoved=Ҫ���Ʋ��˵�ҽ���Ƿ�ת��
'      strAlter=���θ��Ƶ�ҽ����Ҫ�л���Ч��ҽ��ID(��ID):123,456,...
'      ShowMe=Ҫ���Ƶ�ҽ������ID��
    Set mfrmParent = frmParent
    mstrPrivs = strPrivs
    mbln��ʿվ = bln��ʿվ
    mlngǰ��ID = lngǰ��ID
    mlng����ID = lng����ID
    If TypeName(varTime) = "String" Then
        mstr�Һŵ� = varTime
        mlng��ҳID = 0
    Else
        mlng��ҳID = varTime
        mstr�Һŵ� = ""
    End If
    mblnMoved = blnMoved
    strAlter = "": mstrAlter = strAlter
    
    Me.Show 1, frmParent
    
    lng����ID = mlng����ID
    If TypeName(varTime) = "String" Then
        varTime = mstr�Һŵ�
    Else
        varTime = mlng��ҳID
    End If
    blnMoved = mblnMoved
    strAlter = mstrAlter
    ShowMe = mstrIDs
End Function

Private Function LoadPatients() As Boolean
'���ܣ���ȡ����ý�����ͬ��Χ�Ĳ����б�
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer
    Dim lng����ID As Long, intBedLen As Long
        
    On Error GoTo errH
    
    If mlngǰ��ID <> 0 Then
        cmdPati.Visible = False
        If mstr�Һŵ� <> "" Then
            strSQL = "Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,A.����" & _
                " From ������Ϣ A,���˹Һż�¼ B Where A.����ID=B.����ID And A.����ID=[1] And B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        Else
            strSQL = _
                "Select A.����ID,B.��ҳID,A.סԺ��,A.����,A.�Ա�,A.����," & _
                " B.��Ժ����,B.��Ժ����,B.סԺҽʦ,B.��Ժ���� as ����,B.�ѱ�," & _
                " B.����,B.��Ժ����ID as ����ID,B.��ǰ����ID as ����ID,C.���� as ����ȼ�," & _
                " B.״̬,B.����ת��,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ" & _
                " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C" & _
                " Where A.����ID=B.����ID And B.����ȼ�ID=C.ID(+)" & _
                " And A.����ID=[1] And B.��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        End If
    Else
        If mstr�Һŵ� <> "" Then
            '�ṩ��ǰҽ�����ھ���Ĳ����嵥��ѡ��:�������ݲ��漰�жϺͶ�ȡ"H���˹Һż�¼"
            strSQL = "Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,A.����" & _
                " From ������Ϣ A,���˹Һż�¼ B Where A.����ID=B.����ID And A.����ID=[1] And B.NO=[2]"
            strSQL = strSQL & " Union " & _
                " Select B.NO,B.����ID,B.�����,B.����,B.�Ա�,B.����,A.����" & _
                " From ������Ϣ A,���˹Һż�¼ B" & _
                " Where A.����ID=B.����ID And B.ִ��״̬=2 And B.ִ����||''=[3]" & _
                " Order By NO"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, UserInfo.����)
        Else
            strSQL = "Select ��Ժ����ID as ����ID,��ǰ����ID as ����ID From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
                
            '�ṩ��ǰ����/��������Ժ�����嵥��ѡ��
            lng����ID = IIF(mbln��ʿվ, Nvl(rsTmp!����ID, 0), Nvl(rsTmp!����ID, 0))
            intBedLen = GetMaxBedLen(lng����ID, Not mbln��ʿվ)
            strSQL = _
                "Select A.����ID,B.��ҳID,A.סԺ��,A.����,A.�Ա�,A.����,B.��Ժ����,B.��Ժ����," & _
                " B.סԺҽʦ,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,B.����," & _
                " B.��Ժ����ID as ����ID,C.���� as ����ȼ�,B.״̬,B.����ת��," & _
                " Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ" & _
                " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C" & _
                " Where A.����ID=B.����ID And B.����ȼ�ID=C.ID(+) And A.����ID=[1] And B.��ҳID=[2]"
            strSQL = strSQL & " Union " & _
                "Select A.����ID,B.��ҳID,A.סԺ��,A.����,A.�Ա�,A.����,B.��Ժ����,B.��Ժ����," & _
                " B.סԺҽʦ,LPAD(B.��Ժ����," & intBedLen & ",' ') as ����,B.�ѱ�,B.����," & _
                " B.��Ժ����ID as ����ID,C.���� as ����ȼ�,B.״̬,B.����ת��," & _
                " Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ" & _
                " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C" & _
                " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=C.ID(+)" & _
                " And (B.��Ժ����>=Sysdate-30 Or Nvl(B.״̬,0)=3 Or B.��Ժ���� is NULL And Nvl(B.״̬,0)<>3)" & _
                IIF(mbln��ʿվ, " And B.��ǰ����ID=[3]", " And B.��Ժ����ID=[3]") & _
                IIF(Not mbln��ʿվ And InStr(mstrPrivs, "���Ʋ���") = 0, " And B.סԺҽʦ=[4]", "") & _
                " Order by ����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, lng����ID, UserInfo.����)
        End If
    End If
    
    lvwPati.ListItems.Clear
    For i = 1 To rsTmp.RecordCount
        If mstr�Һŵ� <> "" Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!NO, rsTmp!����, , "Pati")
            objItem.SubItems(1) = Nvl(rsTmp!�����)
            objItem.SubItems(2) = Nvl(rsTmp!NO)
            objItem.SubItems(3) = Nvl(rsTmp!�Ա�)
            objItem.SubItems(4) = Nvl(rsTmp!����)
            
            '���ղ����ú�ɫ��ʾ
            If Not IsNull(rsTmp!����) Then
                Call SetItemColor(objItem, vbRed)
            End If
            
            '��ʾ��ʼ���˵���Ϣ
            If rsTmp!����ID = mlng����ID And rsTmp!NO = mstr�Һŵ� Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    lblPati.Caption = "�����:" & .SubItems(1) & "���Һŵ�:" & .SubItems(2) & _
                        "���Ա�:" & .SubItems(3) & "������:" & .SubItems(4)
                    .Selected = True 'һ��Ҫѡ�е�ǰ����
                End With
            End If
        Else
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!��ҳID, rsTmp!����, , "Pati")
            objItem.SubItems(1) = Nvl(rsTmp!סԺ��)
            objItem.SubItems(2) = Nvl(rsTmp!����)
            objItem.SubItems(3) = Nvl(rsTmp!סԺҽʦ)
            objItem.SubItems(4) = Nvl(rsTmp!�Ա�)
            objItem.SubItems(5) = Nvl(rsTmp!����)
            objItem.SubItems(6) = Nvl(rsTmp!�ѱ�)
            objItem.SubItems(7) = Nvl(rsTmp!����ȼ�)
            objItem.SubItems(8) = Format(rsTmp!��Ժ����, "MM-dd HH:mm")
            objItem.SubItems(9) = Format(Nvl(rsTmp!��Ժ����), "MM-dd HH:mm")
            objItem.SubItems(10) = Nvl(rsTmp!ҽ�Ƹ��ʽ)
            objItem.Tag = Nvl(rsTmp!����ת��, 0)
            
            '���ղ����ú�ɫ��ʾ
            If Not IsNull(rsTmp!����) Then
                Call SetItemColor(objItem, vbRed)
            End If
            
            '��ʾ��ʼ���˵���Ϣ
            If rsTmp!����ID = mlng����ID And rsTmp!��ҳID = mlng��ҳID Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    lblPati.Caption = "סԺ��:" & .SubItems(1) & "������:" & .SubItems(2) & _
                        "���Ա�:" & .SubItems(4) & "������:" & .SubItems(5) & _
                        "���ѱ�:" & .SubItems(6) & "�����ʽ:" & .SubItems(10)
                    .Selected = True 'һ��Ҫѡ�е�ǰ����
                End With
            End If
        End If
        rsTmp.MoveNext
    Next
    
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetItemColor(ByVal objItem As ListItem, ByVal lngColor As Long)
    Dim i As Long
    
    objItem.ForeColor = lngColor
    For i = 1 To objItem.ListSubItems.Count
        objItem.ListSubItems(i).ForeColor = lngColor
    Next
End Sub

Private Sub cmdAll_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, colID)) <> 0 And RowCanSelect(i) = 0 Then
                .TextMatrix(i, colѡ��) = -1
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, colѡ��) = 0
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim lng��ID As Long, i As Long
    Dim strIDs As String, strAlter As String
    
    With vsAdvice
        'ȡһ��ҽ����ID
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, colѡ��)) <> 0 Then
                lng��ID = Val(.TextMatrix(i, col���ID))
                If lng��ID = 0 Then lng��ID = Val(.TextMatrix(i, colID))
                
                'ѡ���Ʋ���
                If InStr(strIDs & ",", "," & lng��ID & ",") = 0 Then
                    strIDs = strIDs & "," & lng��ID
                End If
                
                '�л���Ч����
                If .TextMatrix(i, col��Ч) <> .Cell(flexcpData, i, col��Ч) Then
                    If InStr(strAlter & ",", "," & lng��ID & ",") = 0 Then
                        strAlter = strAlter & "," & lng��ID
                    End If
                End If
            End If
        Next
        strAlter = Mid(strAlter, 2)
        strIDs = Mid(strIDs, 2)
        If strIDs = "" Then
            MsgBox "��ѡ��Ҫ���Ƶ�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    mstrAlter = strAlter
    mstrIDs = strIDs
    Unload Me
End Sub

Private Sub cmdPati_Click()
    If mstr�Һŵ� <> "" Then
        lvwPati.ListItems("_" & mlng����ID & "_" & mstr�Һŵ�).Selected = True
    Else
        lvwPati.ListItems("_" & mlng����ID & "_" & mlng��ҳID).Selected = True
    End If
    lvwPati.SelectedItem.EnsureVisible
    lvwPati.Left = txtPati.Left + fraPati.Left
    lvwPati.Top = txtPati.Top + txtPati.Height + fraPati.Top
    lvwPati.Height = vsAdvice.Height - 300
    lvwPati.ZOrder
    lvwPati.Visible = True
    lvwPati.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    ElseIf KeyCode = vbKeyEscape Then
        If lvwPati.Visible Then
            lvwPati.Visible = False
        Else
            Unload Me
        End If
    ElseIf KeyCode = vbKeyF4 Or KeyCode = vbKeyDown Then
        If Not (KeyCode = vbKeyDown And Shift <> vbAltMask) Then
            If Me.ActiveControl Is txtPati Then
                If cmdPati.Visible And cmdPati.Enabled Then cmdPati_Click
            End If
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Dim strLvw As String
    If mstr�Һŵ� <> "" Then
        strLvw = "����,1000,0,1;�����,1000,0,1;�Һŵ�,1000,0,1;�Ա�,600,0,1;����,600,0,1"
    Else
        strLvw = "����,1000,0,1;סԺ��,1000,0,1;����,630,0,1;סԺҽʦ,1000,0,1;�Ա�,600,0,1;����,600,0,1;�ѱ�,850,0,1;����ȼ�,1150,0,1;��Ժ����,1180,0,1;��Ժ����,1180,0,1;���ʽ,1500,0,1"
    End If
    Call zlControl.LvwSelectColumns(lvwPati, strLvw, True)
    Call RestoreWinState(Me, App.ProductName, IIF(mstr�Һŵ� <> "", 1, 2))
    If mlng��ҳID <> 0 Then
        vsAdvice.FrozenCols = col��Ч + 1
    Else
        vsAdvice.FrozenCols = colѡ�� + 1
    End If
    
    Call LoadPatients
    Call LoadAdvice
    
    mstrIDs = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraPati.Top = -75
    fraPati.Left = 0
    fraPati.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraPati.Top + fraPati.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - vsAdvice.Top - cmdOK.Height * 1.6
        
    cmdHelp.Top = Me.ScaleHeight - cmdAll.Height * 1.3
    cmdAll.Top = cmdHelp.Top
    cmdClear.Top = cmdAll.Top
    cmdOK.Top = cmdAll.Top
    cmdCancel.Top = cmdAll.Top
    
    If Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3) < 5000 Then
        cmdCancel.Left = 5000
    Else
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3)
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, IIF(mstr�Һŵ� <> "", 1, 2))
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    If mblnItem Then Call lvwPati_KeyPress(13)
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Dim lng����ID As Long, lng��ҳID As Long, strNO As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwPati.SelectedItem Is Nothing Then
            lng����ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(0))
            If mstr�Һŵ� <> "" Then
                strNO = Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1)
                If lng����ID = mlng����ID And strNO = mstr�Һŵ� Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng����ID = lng����ID
                    mstr�Һŵ� = strNO
                    mblnMoved = MovedByNO(strNO, "���˹Һż�¼")
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                    lblPati.Caption = "�����:" & .SubItems(1) & "���Һŵ�:" & .SubItems(2) & _
                        "���Ա�:" & .SubItems(3) & "������:" & .SubItems(4)
                End With
            Else
                lng��ҳID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1))
                If lng����ID = mlng����ID And lng��ҳID = mlng��ҳID Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng����ID = lng����ID
                    mlng��ҳID = lng��ҳID
                    mblnMoved = Val(.Tag) = 1
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                    lblPati.Caption = "סԺ��:" & .SubItems(1) & "������:" & .SubItems(2) & _
                        "���Ա�:" & .SubItems(4) & "������:" & .SubItems(5) & _
                        "���ѱ�:" & .SubItems(6) & "  ���ʽ:" & .SubItems(10)
                End With
            End If
            lvwPati.Visible = False
            
            '��ȡ����ʾ����ҽ��
            Call LoadAdvice
            
            vsAdvice.SetFocus
        End If
    End If
End Sub

Private Sub lvwPati_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnItem = False
End Sub

Private Sub lvwPati_Validate(Cancel As Boolean)
    lvwPati.Visible = False
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub RowSelectSame(ByVal lngRow As Long)
'���ܣ�����ָ����(����Ϊ������)��ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                    .TextMatrix(i, col��Ч) = .TextMatrix(lngRow, col��Ч)
                    .Cell(flexcpFontBold, i, col��Ч) = .Cell(flexcpFontBold, lngRow, col��Ч)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                    .TextMatrix(i, col��Ч) = .TextMatrix(lngRow, col��Ч)
                    .Cell(flexcpFontBold, i, col��Ч) = .Cell(flexcpFontBold, lngRow, col��Ч)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                    .TextMatrix(i, col��Ч) = .TextMatrix(lngRow, col��Ч)
                    .Cell(flexcpFontBold, i, col��Ч) = .Cell(flexcpFontBold, lngRow, col��Ч)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                    .TextMatrix(i, col��Ч) = .TextMatrix(lngRow, col��Ч)
                    .Cell(flexcpFontBold, i, col��Ч) = .Cell(flexcpFontBold, lngRow, col��Ч)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Function LoadAdvice() As Boolean
'���ܣ���ȡ��ǰ����ָ����ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    On Error GoTo errH
    
    '�ſ������Ͳ������ڵ�����
    strSQL = "Select Distinct A.ID,A.���,A.���ID,A.ҽ����Ч,A.��ʼִ��ʱ��,A.������ĿID," & _
        " A.ҽ������,A.��������,A.ִ��Ƶ��,A.ҽ������,C.���� as ִ�п���,A.ִ��ʱ�䷽��,A.�շ�ϸĿID," & _
        " A.�걾��λ,B.���,B.����,B.���㵥λ,A.�ܸ����� as ����,E.�����װ,E.���ﵥλ,E.סԺ��װ,E.סԺ��λ," & _
        " B.����ʱ��,B.�������,D.����ʱ�� as �շѳ���,D.������� as �շѷ���" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D,ҩƷ��� E" & _
        " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=C.ID(+) And A.�շ�ϸĿID=D.ID(+) And A.�շ�ϸĿID=E.ҩƷID(+)" & _
        " And A.ҽ��״̬ Not IN(2,4) And A.��ʼִ��ʱ�� is Not Null And A.������Դ<>3" & _
        IIF(mstr�Һŵ� <> "", " And A.����ID+0=[1] And A.�Һŵ�=[2]", " And A.����ID=[1] And A.��ҳID=[2]") & _
        " Order by A.���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, IIF(mstr�Һŵ� <> "", mstr�Һŵ�, mlng��ҳID))
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '����������
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, colѡ��) = 0
                .TextMatrix(i, colID) = rsTmp!ID
                .TextMatrix(i, col���ID) = Nvl(rsTmp!���ID)
                .TextMatrix(i, col�������) = Nvl(rsTmp!���, "*")
                .TextMatrix(i, col������ĿID) = Nvl(rsTmp!������ĿID)
                .TextMatrix(i, col�շ�ϸĿID) = Nvl(rsTmp!�շ�ϸĿID)
                .TextMatrix(i, col��Ч) = IIF(Nvl(rsTmp!ҽ����Ч, 0) = 0, "����", "����")
                .Cell(flexcpData, i, col��Ч) = .TextMatrix(i, col��Ч)
                .TextMatrix(i, colʱ��) = Format(rsTmp!��ʼִ��ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, col����) = rsTmp!ҽ������
                .Cell(flexcpData, i, col����) = Nvl(rsTmp!�걾��λ) '����걾
                .TextMatrix(i, col����) = FormatEx(Nvl(rsTmp!��������), 4)
                If Not IsNull(rsTmp!��������) Then
                    .TextMatrix(i, col������λ) = Nvl(rsTmp!���㵥λ)
                End If
                If InStr(",5,6,", Nvl(rsTmp!���, "*")) > 0 Then
                    If mstr�Һŵ� <> "" Then
                        If Not IsNull(rsTmp!����) And Not IsNull(rsTmp!�����װ) Then
                            .TextMatrix(i, col����) = FormatEx(rsTmp!���� / rsTmp!�����װ, 5)
                        End If
                        If Nvl(rsTmp!ҽ����Ч, 0) = 1 Then
                            .TextMatrix(i, col������λ) = Nvl(rsTmp!���ﵥλ)
                        End If
                    Else
                        If Not IsNull(rsTmp!����) And Not IsNull(rsTmp!סԺ��װ) Then
                            .TextMatrix(i, col����) = FormatEx(rsTmp!���� / rsTmp!סԺ��װ, 5)
                        End If
                        If Nvl(rsTmp!ҽ����Ч, 0) = 1 Then
                            .TextMatrix(i, col������λ) = Nvl(rsTmp!סԺ��λ)
                        End If
                    End If
                Else
                    If Not IsNull(rsTmp!����) Then
                        .TextMatrix(i, col����) = rsTmp!����
                    End If
                    If Nvl(rsTmp!ҽ����Ч, 0) = 1 Then
                        .TextMatrix(i, col������λ) = Nvl(rsTmp!���㵥λ)
                    End If
                End If
                
                .TextMatrix(i, colƵ��) = Nvl(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, col����) = Nvl(rsTmp!ҽ������)
                .TextMatrix(i, colִ��ʱ��) = Nvl(rsTmp!ִ��ʱ�䷽��)
                .TextMatrix(i, colִ�п���) = Nvl(rsTmp!ִ�п���)
                
                '���������ؼ��÷���ʾ
                If InStr(",C,D,F,G,E,", Nvl(rsTmp!���, "*")) > 0 And Not IsNull(rsTmp!���ID) Then
                    .RowHidden(i) = True
                ElseIf Nvl(rsTmp!���) = "7" Then
                    .RowHidden(i) = True
                ElseIf Nvl(rsTmp!���) = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col���ID)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col�������)) > 0 Then
                    '��ҩ;��
                    .RowHidden(i) = True
                    '��ʾ��ҩ;��
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���ID)) = rsTmp!ID Then
                            .TextMatrix(j, col�÷�) = rsTmp!����
                        Else
                            Exit For
                        End If
                    Next
                ElseIf Nvl(rsTmp!���) = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col���ID)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col�������)) > 0 Then
                    '��ҩ�÷������ɼ�����
                    .TextMatrix(i, col�÷�) = rsTmp!����
                    
                    '��ҩ������ִ�п���
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���ID)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col�������)) > 0 Then
                                .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '��ҩ����
                    If .TextMatrix(i - 1, col�������) <> "C" Then
                        .TextMatrix(i, col������λ) = "��"
                    End If
                End If
                
                '��ǰ������г����򲻷������Ŀ
                If Not IsNull(rsTmp!������ĿID) Then
                    If Not (IsNull(rsTmp!����ʱ��) Or Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = 1
                    ElseIf Not (Nvl(rsTmp!�������, 0) = 3 Or Nvl(rsTmp!�������, 0) = IIF(mstr�Һŵ� <> "", 1, 2)) Then
                        .RowData(i) = 1
                    ElseIf Not IsNull(rsTmp!�շ�ϸĿID) Then
                        '��ҩƷ,ͬʱҪ�жϵ��շ���ĿĿ¼
                        If Not (IsNull(rsTmp!�շѳ���) Or Format(Nvl(rsTmp!�շѳ���), "yyyy-MM-dd") = "3000-01-01") Then
                            .RowData(i) = 1
                        ElseIf Not (Nvl(rsTmp!�շѷ���, 0) = 3 Or Nvl(rsTmp!�շѷ���, 0) = IIF(mstr�Һŵ� <> "", 1, 2)) Then
                            .RowData(i) = 1
                        End If
                    End If
                End If

                rsTmp.MoveNext
            Next
        End If
        If mlng��ҳID <> 0 Then
            .Cell(flexcpBackColor, .FixedRows, colѡ��, .Rows - 1, col��Ч) = &HC0FFC0
        End If
        
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col����
        .ColHidden(col��Ч) = mstr�Һŵ� <> ""
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RowCanSelect(ByVal lngRow As Long) As Long
'���ܣ��ж�ָ���е�(���)ҽ���ɷ�ѡ��
'���أ��������ѡ�񣬷���0,���򷵻��к�
    Dim i As Long
    
    With vsAdvice
        If .RowData(lngRow) = 1 Then RowCanSelect = lngRow: Exit Function
        
        If Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) _
                    Or Val(.TextMatrix(i, colID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colID)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = colѡ�� Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col���� Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colѡ�� Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = colʱ��
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = colƵ��: lngRight = col�÷�
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            If Col = lngLeft And lngLeft = col��Ч Then
                SetBkColor hDC, SysColor2RGB(.Cell(flexcpBackColor, Row, lngLeft))
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> colѡ�� Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, colID)) = 0 Then Exit Sub
                
                '����Ƿ���Ա�ѡ��
                i = RowCanSelect(.Row)
                If i > 0 And Val(.TextMatrix(.Row, colѡ��)) = 0 Then
                    MsgBox "��Ϊҽ��""" & .TextMatrix(i, col����) & """��Ӧ����Ŀ�ѳ�����������ƥ�䣬��ҽ�����ܱ�ѡ��", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                If .Col = col��Ч And mlng��ҳID <> 0 Then
                    If CanAlterType(.Row) Then
                        .TextMatrix(.Row, .Col) = IIF(.TextMatrix(.Row, .Col) = "����", "����", "����")
                        .Cell(flexcpFontBold, .Row, .Col) = .TextMatrix(.Row, .Col) <> .Cell(flexcpData, .Row, .Col)
                        If .Cell(flexcpFontBold, .Row, .Col) Then
                            .TextMatrix(.Row, colѡ��) = -1
                        End If
                        Call RowSelectSame(.Row)
                    End If
                Else
                    .TextMatrix(.Row, colѡ��) = IIF(Val(.TextMatrix(.Row, colѡ��)) = 0, -1, 0)
                    Call RowSelectSame(.Row)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    If Col <> colѡ�� Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(vsAdvice.Row, colID)) = 0 Then
        Cancel = True
    Else
        i = RowCanSelect(Row)
        If i > 0 Then
            Cancel = True
            MsgBox "��Ϊҽ��""" & vsAdvice.TextMatrix(i, col����) & """��Ӧ����Ŀ�ѳ�����������ƥ�䣬��ҽ�����ܱ�ѡ��", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function CanAlterType(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ����ҽ���Ƿ�����л���Ч
'������lngRow=�ɼ���ҽ����
'˵���������л���Ч��������
'   1.�ɳ�����ִ��Ƶ��=0(��ѡƵ��),2(������)
'   2.��������ִ��Ƶ��=0(��ѡƵ��),1(һ����);ҩƷ����ָ���˹��
    Dim rsMore As New ADODB.Recordset
    Dim strSQL As String, strType As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, colID)) = 0 Then
            CanAlterType = True: Exit Function
        ElseIf Val(.TextMatrix(lngRow, col������ĿID)) = 0 Then
            '��������Ŀ����л�
            CanAlterType = True: Exit Function
        ElseIf RowIn�䷽��(lngRow) Then
            '��ҩ�䷽�̶������л�
            CanAlterType = True: Exit Function
        ElseIf RowIn������(lngRow) Then
            '�����Լ�����Ϊ׼�ж�
            lngRow = .FindRow(.TextMatrix(lngRow, colID), , col���ID)
            If lngRow = -1 Then Exit Function
        End If
    
        strType = IIF(.TextMatrix(lngRow, col��Ч) = "����", "����", "����")
        
        '��ԭʼƵ��Ϊ׼�ж�:��Ϊ��ѡ��Ƶ�ʵĿ�����ȱ��һ����
        strSQL = "Select ִ��Ƶ�� From ������ĿĿ¼ Where ID=[1]"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, col������ĿID)))
        
        If strType = "����" Then
            If InStr(",0,2,", Nvl(rsMore!ִ��Ƶ��, 0)) = 0 Then Exit Function
        Else
            If InStr(",0,1,", Nvl(rsMore!ִ��Ƶ��, 0)) = 0 Then Exit Function
            If InStr(",5,6,", .TextMatrix(lngRow, col�������)) > 0 Then
                Call GetRowScope(lngRow, lngBegin, lngEnd)
                For i = lngBegin To lngEnd
                    If InStr(",5,6,", .TextMatrix(i, col�������)) > 0 Then
                        If Val(.TextMatrix(i, col�շ�ϸĿID)) = 0 Then Exit Function
                    End If
                Next
            End If
        End If
    End With
    CanAlterType = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'���ܣ���ȡ��ID��ͬ��һ��ҽ���кŷ�Χ(ע�⿼��һ����ҩ�еĿ���)
    Dim lngS��ID As Long, lngO��ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS��ID = IIF(Val(.TextMatrix(lngRow, col���ID)) = 0, Val(.TextMatrix(lngRow, colID)), Val(.TextMatrix(lngRow, col���ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO��ID = IIF(Val(.TextMatrix(i, col���ID)) = 0, Val(.TextMatrix(i, colID)), Val(.TextMatrix(i, col���ID)))
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '��������
                If lngO��ID = lngS��ID Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO��ID = IIF(Val(.TextMatrix(i, col���ID)) = 0, Val(.TextMatrix(i, colID)), Val(.TextMatrix(i, col���ID)))
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '��������
                If lngO��ID = lngS��ID Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function RowIn������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����ڼ�������е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "E" And Val(.TextMatrix(lngRow, col���ID)) = 0 Then
            '�ɼ�������
            If .TextMatrix(lngRow - 1, col�������) = "C" _
                And Val(.TextMatrix(lngRow - 1, col���ID)) = .TextMatrix(lngRow, colID) Then
                RowIn������ = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, col�������) = "C" And Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
            '������Ŀ��
            RowIn������ = True: Exit Function
        End If
    End With
End Function

Private Function RowIn�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�������ҩ�䷽�е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "E" Then
            If Val(.TextMatrix(lngRow, col���ID)) = 0 Then
                '�÷���
                If Val(.TextMatrix(lngRow - 1, col���ID)) = .TextMatrix(lngRow, colID) _
                    And .TextMatrix(lngRow - 1, col�������) = "E" Then
                    RowIn�䷽�� = True: Exit Function
                End If
            Else
                '�巨��
                If .TextMatrix(lngRow - 1, col�������) = "7" _
                    And Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    RowIn�䷽�� = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, col�������) = "7" And Val(.TextMatrix(lngRow, col���ID)) <> 0 Then
            '��ҩ��
            RowIn�䷽�� = True: Exit Function
        End If
    End With
End Function
