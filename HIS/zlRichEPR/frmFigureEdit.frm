VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMarkMapEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������ͼ�α༭"
   ClientHeight    =   3300
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7305
   Icon            =   "frmFigureEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkAspectRatio 
      Caption         =   "�����ݺ��(&M)"
      Height          =   255
      Left            =   1305
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2430
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtHeight 
      Height          =   300
      Left            =   2340
      TabIndex        =   5
      Top             =   2070
      Width           =   795
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   1290
      TabIndex        =   4
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   3870
      TabIndex        =   3
      Top             =   1650
      Width           =   375
   End
   Begin VB.CheckBox chkFitMode 
      Alignment       =   1  'Right Justify
      Caption         =   "�ʺϴ�С(&F)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5985
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2970
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1260
      Top             =   2745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"frmFigureEdit.frx":058A
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2025
      TabIndex        =   7
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3150
      TabIndex        =   8
      Top             =   2850
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1290
      TabIndex        =   2
      Top             =   1650
      Width           =   1350
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Top             =   1230
      Width           =   2955
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1290
      TabIndex        =   0
      Top             =   825
      Width           =   795
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   570
      TabIndex        =   13
      Top             =   600
      Width           =   3735
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   12
      Top             =   2745
      Width           =   4320
   End
   Begin zlRichEPR.ucCanvas Canvas 
      Height          =   2805
      Left            =   4410
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   90
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4948
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3195
      TabIndex        =   21
      Top             =   2130
      Width           =   360
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2115
      TabIndex        =   20
      Top             =   2130
      Width           =   180
   End
   Begin VB.Label lblSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��С(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   19
      Top             =   2130
      Width           =   630
   End
   Begin VB.Label lblPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͼƬ:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3330
      TabIndex        =   18
      Top             =   1710
      Width           =   450
   End
   Begin VB.Label lblColor 
      Caption         =   "��ɫ���:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4410
      TabIndex        =   11
      Top             =   2970
      Width           =   2190
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   17
      Top             =   1710
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   45
      Picture         =   "frmFigureEdit.frx":0647
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "�༭��Ӧ�õ�ͳһ�Ĳ������ͼ����Դ�������������������Ա����ʹ�á�"
      Height          =   345
      Left            =   585
      TabIndex        =   16
      Top             =   135
      Width           =   3660
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   15
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   585
      TabIndex        =   14
      Top             =   885
      Width           =   630
   End
End
Attribute VB_Name = "frmMarkMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'���룺
'   1���ϼ�����ͨ��������ShowMe�������������塢�༭����ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private mstrItemCode As String      '���༭����Ŀ���룬�޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0��
Private mblnOK As Boolean           '�Ƿ���ɱ༭�˳�

'��ʱ����
Dim rsTemp As New ADODB.Recordset

'################################################################################################################
'-- λͼ����
Private WithEvents DIBFilter As cDIBFilter      ' DIB �˾�����(24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Private WithEvents DIBDither As cDIBDither      ' DIB ��������(1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Private DIBPal               As New cDIBPal     ' DIB ��ɫ����� (1, 4, 8 bpp)
Private DIBSave              As New cDIBSave    ' Save ���� (BMP)  (1, 4, 8, 24 bpp)
Private DIBbpp               As Byte            ' ��ǰ��ɫ���
Private WithEvents cPicEditor As cPictureEditor     ' ͼƬ�༭����
Attribute cPicEditor.VB_VarHelpID = -1
Private m_LastFilename As String                    ' ���򿪵�ͼƬλ��
Private m_Temp As String                            ' ��ʱ�ļ�·��
Private m_AppID As Long
'-- GDI+
Private m_GDIpToken         As Long         ' ���ڹر� GDI+
Private mblnAdd As Boolean                  ' �Ƿ�������
Private Const MAX_PIXELS_SIZE As Long = 4000000
Private W As Long, chgW As Boolean
Private H As Long, chgH As Boolean
Private mfrmParent As Object

Public Function ShowMe(ByRef frmParent As Object, ByVal blnAdd As Boolean, Optional ByVal strItemCode As String, _
    Optional oDIB As cDIB) As String
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '���أ�ȷ�������������޸ĵı��룻ȡ������""
    '---------------------------------------------------
    mblnAdd = blnAdd
    If mblnAdd Then
        Me.Tag = "����"
    Else
        Me.Tag = "�޸�"
    End If
    
    Set mfrmParent = frmParent
    
    If mblnAdd = False Then
        Set Canvas.DIB = oDIB
        Me.Canvas.Resize
        If Me.Canvas.DIB.hDIB <> 0 Then
            W = oDIB.Width
            H = oDIB.Height
            txtWidth = W
            txtHeight = H
            txtWidth.Enabled = True
            txtHeight.Enabled = True
            chkAspectRatio.Enabled = True
            chkFitMode.Enabled = True
            lblSize.Enabled = True
            lblColor.Enabled = True
            lblColor = "��ɫ���:24 λ"
        Else
            txtWidth.Enabled = False
            txtHeight.Enabled = False
            chkAspectRatio.Enabled = False
            chkFitMode.Enabled = False
            lblSize.Enabled = False
            lblColor.Enabled = False
        End If
    End If
    
    mstrItemCode = strItemCode
    
    '��ȡ��Ϣ
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ����,����,���� From �������ͼ�� Where ����=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mstrItemCode)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt����.Text = !����: Me.txt����.Text = !����
            Me.txt����.Text = IIf(IsNull(!����), "", !����)
        End If
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
    End With
    If Me.Tag = "����" Then
        gstrSQL = "Select nvl(max(����),'" & String(Me.txt����.MaxLength, "0") & "') as ���� From �������ͼ��"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
        Me.txt����.Text = Format(Val(rsTemp!����) + 1, String(Me.txt����.MaxLength, "0"))
    End If
    
    txtWidth.Enabled = (Me.Canvas.DIB.hDIB <> 0)
    txtHeight.Enabled = (Me.Canvas.DIB.hDIB <> 0)
    '��ʾ����
    Me.Show vbModal, frmParent
    If mblnOK Then
        ShowMe = Trim(Me.txt����.Text)
        Set frmParent.Canvas.DIB = Me.Canvas.DIB
        frmParent.Canvas.Resize
    Else
        ShowMe = ""
    End If
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = ""
End Function

Private Sub chkFitMode_Click()
    Canvas.FitMode = CBool(chkFitMode)
    Call Canvas.Resize
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim arySql() As String, lngSql As Long

    If Trim(Me.txt����.Text) = "" Then MsgBox "��������룡", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If Len(Me.txt����.Text) < Me.txt����.MaxLength Then MsgBox "���볤�Ȳ��㣡", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If Trim(Me.txt����.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    End If
    
    '���ݱ���
    gstrSQL = "'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
    If Me.Tag = "����" Then
        If Me.Canvas.DIB.hDIB = 0 Then
            MsgBox "����ѡ��һ��ͼƬ��", vbOKOnly + vbInformation
            If cmdSelect.Visible And cmdSelect.Enabled Then cmdSelect.SetFocus
            Exit Sub
        Else
            gstrSQL = "ZL_�������ͼ��_INSERT(" & gstrSQL & ")"
        End If
    Else
        '�޸�ģʽ��Ҫȷ���ò���ͼ�λ�����
        gstrSQL = "select count(*) from �������ͼ�� where ���� =[1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mstrItemCode)
        If rsTemp(0) = 0 Then
            MsgBox "��ͼƬ�Ѿ��������û�ɾ��������ʧ�ܣ�", vbOKOnly + vbInformation, "����ʧ��"
            Unload Me
            mfrmParent.zlRefLists
            Exit Sub
        End If
        rsTemp.Close
        gstrSQL = "ZL_�������ͼ��_UPDATE('" & mstrItemCode & "'," & gstrSQL & ")"
    End If
        
    ReDim Preserve arySql(0 To 0)
    arySql(0) = gstrSQL

    '����ͼƬ�³ߴ�
    If Me.Canvas.DIB.hDIB <> 0 Then
        If (txtWidth * txtHeight > MAX_PIXELS_SIZE) Then
            Call MsgBox(vbCrLf & _
                "ͼƬ��С�����������Χ(4M ����)" & vbCrLf & vbCrLf & _
                "���СͼƬ�ߴ磡", vbExclamation)
            txtWidth.SelStart = 0: txtWidth.SelLength = Len(txtWidth): txtWidth.SetFocus
            Exit Sub
        End If
    
        Dim lngR As Long, strMsg As String
        If (txtWidth <> Me.Canvas.DIB.Width) Or (txtHeight <> Me.Canvas.DIB.Height) Then
            If txtWidth < 10 Or txtHeight < 10 Then
                lngR = MsgBox("ע�⣺ͼƬ�ߴ��С��һ�����ص���10��10��ͼƬ���޷���Ч���ã�" & vbCrLf & _
                    "�Ƿ������ ѡ���ǡ�������ѡ����ȡ����", vbYesNo + vbQuestion, gstrSysName)
            ElseIf txtWidth / Me.Canvas.DIB.Width < 0.5 And txtHeight / Me.Canvas.DIB.Height < 0.5 Then
                lngR = MsgBox("ע�⣺ͼƬ�ߴ�С��ԭʼ�ߴ��һ�룬�⽫����ͼƬ������ʧ�����Ҳ��ɻָ���" & vbCrLf & _
                    "�Ƿ������ ѡ���ǡ�������ѡ����ȡ����", vbYesNo + vbQuestion, gstrSysName)
            ElseIf txtWidth / Me.Canvas.DIB.Width < 0.5 Then
                lngR = MsgBox("ע�⣺ͼƬ���С��ԭʼ�ߴ��һ�룬�⽫����ͼƬ������ʧ�����Ҳ��ɻָ���" & vbCrLf & _
                    "�Ƿ������ ѡ���ǡ�������ѡ����ȡ����", vbYesNo + vbQuestion, gstrSysName)
            ElseIf txtHeight / Me.Canvas.DIB.Height < 0.5 Then
                lngR = MsgBox("ע�⣺ͼƬ�߶�С��ԭʼ�ߴ��һ�룬�⽫����ͼƬ������ʧ�����Ҳ��ɻָ���" & vbCrLf & _
                    "�Ƿ������ ѡ���ǡ�������ѡ����ȡ����", vbYesNo + vbQuestion, gstrSysName)
            Else
                lngR = MsgBox("ע�⣺�ı�ͼ��ߴ罫����ʧͼƬ���������Ҳ��ɻָ���" & vbCrLf & _
                    "�Ƿ������ ѡ���ǡ�������ѡ����ȡ����", vbYesNo + vbQuestion, gstrSysName)
            End If
            If lngR = vbYes Then
                Screen.MousePointer = vbHourglass
                Call mGdIpEx.ScaleDIB(Me.Canvas.DIB, txtWidth, txtHeight, True)
                Call Me.Canvas.RemoveCropRectangle
                Call Me.Canvas.Resize
                Screen.MousePointer = vbNormal
            Else
                txtWidth = Me.Canvas.DIB.Width
                txtHeight = Me.Canvas.DIB.Height
                Exit Sub
            End If
        End If
    
        'ͬʱ����λͼ
        Dim strFileName As String
        Screen.MousePointer = vbHourglass
        strFileName = m_Temp & "\R" & m_AppID & ".jpg"
        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName, [ImageJPEG], 90)         '90%��ͼƬ����������ѹ��
        
        If gobjFSO.FileExists(strFileName) Then
            If zlBlobSql(0, Trim(Me.txt����.Text), strFileName, arySql) = False Then
                Screen.MousePointer = vbNormal
                MsgBox "���ͼ�α���ʧ��", vbExclamation, gstrSysName
                Exit Sub
            End If
            gobjFSO.DeleteFile strFileName  'ɾ����ʱ�ļ�
        End If
        Screen.MousePointer = vbNormal
    End If
    
    'ִ�б���
    Err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    For lngSql = LBound(arySql) To UBound(arySql)
        Call SQLTest(App.ProductName, Me.Caption, arySql(lngSql))
        gcnOracle.Execute arySql(lngSql), , adCmdStoredProc
        Call SQLTest
    Next
    gcnOracle.CommitTrans
    
    mblnOK = True: Me.Hide
    Exit Sub

errHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    Dim strFileName As String, bSuccess As Boolean, strTmp As String
    dlgThis.InitDir = m_LastFilename
    dlgThis.CancelError = True
    On Error GoTo LL
    dlgThis.ShowOpen
    
    strFileName = dlgThis.Filename
    If Trim(strFileName) <> "" Then
        '-- Create DIB
'        DoEvents
        Call pvSetDIBPicture(pvGetStdPicture(strFileName, bSuccess))
        
        If (bSuccess) Then
            m_LastFilename = strFileName
            W = Me.Canvas.DIB.Width
            H = Me.Canvas.DIB.Height
            txtWidth = W
            txtHeight = H
            lblColor = "��ɫ���:" & DIBbpp & " λ"
            txtWidth.Enabled = True
            txtHeight.Enabled = True
            chkAspectRatio.Enabled = True
            chkFitMode.Enabled = True
            lblSize.Enabled = True
            lblColor.Enabled = True
        End If
    End If
    txtWidth.Enabled = (Me.Canvas.DIB.hDIB <> 0)
    txtHeight.Enabled = (Me.Canvas.DIB.hDIB <> 0)
LL:
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

'################################################################################################################
'## ���ܣ�  ������غ���
'################################################################################################################
Private Function pvGetStdPicture(ByVal sFileName As String, bSuccess As Boolean) As StdPicture
    On Error Resume Next
    If (pvGetExt(sFileName) = "png" Or pvGetExt(sFileName) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGdIpEx.LoadPictureEx(sFileName)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFileName)
    End If
    
    '-- Is there an image ?
    bSuccess = Not (pvGetStdPicture Is Nothing)
    
    If (bSuccess = False) Then
        '-- Nothing loaded
        Call MsgBox("����ͼƬʱ�����������", vbExclamation)
    End If

    On Error GoTo 0
End Function
    
Private Sub pvSetDIBPicture(Image As StdPicture)
  Static lstW As Long
  Static lstH As Long

    If (Not Picture Is Nothing) Then

        '-- Save last DIB dimensions
        lstW = Canvas.DIB.Width
        lstH = Canvas.DIB.Height
        
        '-- Clear palette
        Call DIBPal.Clear
        
        DIBbpp = Canvas.DIB.CreateFromStdPicture(Image, DIBPal, DIBDither)
        
        '-- Select current depth mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Remove Crop rectangle and resize canvas
        Call Canvas.RemoveCropRectangle
        With Canvas.DIB
            If (lstW <> .Width Or lstH <> .Height) Then
                Call Canvas.Resize
              Else
                Call Canvas.Repaint
            End If
        End With
    End If
End Sub

Private Sub pvSetPalMode(ByVal bpp As Long)
  Dim lIdxNew As Long
  Dim lIdxOld As Long
    
    Select Case bpp
        Case 1  '-- 2 colors / Black and White
            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
        Case 4  '-- 16 colors / 16 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
        Case 8  '-- 256 colors / 256 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
        Case 24 '-- True color
            lIdxNew = 8
        Case Else
            Exit Sub
    End Select
End Sub

Private Function pvGetExt(ByVal sFileName As String) As String
    pvGetExt = Mid(sFileName, Len(sFileName) - 2)
End Function

Private Sub Form_Load()
    '-----------------------------------------------------
    m_LastFilename = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", App.Path)
    Dim GpInput As GdiplusStartupInput
    '-- ���� GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("���� GDI+ �����޷�����ͼƬ���룡���� GDI+ DLL �Ƿ���ڻ����𻵣�", vbInformation + vbOKOnly)
        Call Unload(Me)
        Exit Sub
    End If
    
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    m_AppID = Me.hWnd
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    Set cPicEditor = New cPictureEditor
    
    Canvas.FitMode = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", m_LastFilename
    If mblnAdd Then Me.Canvas.DIB.Destroy    '�޸�ģʽ�ǲ���ɾ����DIB�ģ�
    
    LockWindowUpdate 0
    UpdateWindow Me.hWnd
    ' Unload the GDI+ Dll
    Call mGdIpEx.GdiplusShutdown(m_GDIpToken)

    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    Set cPicEditor = Nothing
End Sub

Private Sub txtHeight_Change()
    txtHeight = Val(txtHeight)
    If (Val(txtHeight) = 0) Then
        If Me.Canvas.DIB.hDIB <> 0 Then
            txtHeight = "1"
        Else
            txtHeight = "0"
        End If
        txtHeight.SelLength = 1
    End If
    If (chkAspectRatio) Then
        If (Not chgW) Then
            chgH = True
            If Me.Canvas.DIB.hDIB <> 0 Then
                txtWidth = CInt(txtHeight / H * W)
            Else
                txtWidth = "0"
            End If
            chgH = False
        End If
    End If
End Sub

Private Sub txtHeight_GotFocus()
    txtHeight.SelStart = Len(txtHeight)
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    txtWidth = Val(txtWidth)
    If (Val(txtWidth) = 0) Then
        If Me.Canvas.DIB.hDIB <> 0 Then
            txtWidth = "1"
        Else
            txtWidth = "0"
        End If
        txtWidth.SelLength = 1
    End If
    If (chkAspectRatio) Then
        If (Not chgH) Then
            chgW = True
            If Me.Canvas.DIB.hDIB <> 0 Then
                txtHeight = CInt(txtWidth / W * H)
            Else
                txtHeight = "0"
            End If
            chgW = False
        End If
    End If
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = Len(txtWidth)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
    If Me.Tag = "����" Then
        Me.txt����.Text = zlGetSymbol(Me.txt����.Text, 0)
    End If
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.txt����.Text = zlGetSymbol(Me.txt����.Text, 0)
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
