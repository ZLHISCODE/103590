VERSION 5.00
Begin VB.Form frmCertPicture 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���֤��ͼƬ"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4035
   Icon            =   "frmCertPicture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4035
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cmdDown 
      Appearance      =   0  'Flat
      Caption         =   "��һ��"
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Appearance      =   0  'Flat
      Caption         =   "��һ��"
      Height          =   300
      Left            =   840
      TabIndex        =   0
      Top             =   3720
      Width           =   1100
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��ע��"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3060
      Width           =   615
   End
   Begin VB.Image imgPicture 
      Appearance      =   0  'Flat
      Height          =   2475
      Left            =   840
      Picture         =   "frmCertPicture.frx":6852
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image imgLoad 
      Height          =   375
      Left            =   5160
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmCertPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mlng֤��id As Long
Private mlng��� As Long
Private mstrNote As String
Private mintType As Integer
Private mstrFile As String
Private mlngFirst As Long
Private mlngLeft As Long, mlngTop As Long, mlngHeight As Long
Private mrsMainInfo As ADODB.Recordset
Private mlngCount As Long

Private Sub cmdDown_Click()
    Dim lngTmp As Long
    Dim strTmp As Variant
    Dim strFile As String
    Dim strMainInfo As String
    
    If mlng��� >= mlngCount Then
        MsgBox "���Ѿ������һ��ͼƬ�ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    mlng��� = mlng��� + 1
    Do While Not mlng��� > mlngCount
        mrsMainInfo.Filter = "Index=" & mlng���
        If mrsMainInfo.EOF Then
            Screen.MousePointer = 11
            Call ReadPatPricture(mlng֤��id & "," & mlng���, imgLoad, strFile)
            Screen.MousePointer = 0
            mstrNote = GetImageNote(mlng֤��id, mlng���)
            If imgLoad.Picture <> 0 Then
                lngTmp = mrsMainInfo.RecordCount + 1
                If imgLoad.Picture <> 0 Then
                    strMainInfo = strFile & "|" & mstrNote
                    mrsMainInfo.AddNew Array("���", "Index", "��Ϣֵ"), Array(lngTmp, mlng���, strMainInfo)
                End If
            End If
        End If
        mrsMainInfo.Filter = "Index=" & mlng���
        If Not mrsMainInfo.EOF Then
            strTmp = Split(mrsMainInfo!��Ϣֵ, "|")
            txtNote = strTmp(1)
            imgPicture.Picture = LoadPicture(strTmp(0))
            Exit Sub
        End If
        mlng��� = mlng��� + 1
   Loop
End Sub

Private Sub InitBaseInfo()
    Dim arrMainFileds() As Variant

    '��ʼ����¼��
    '1����¼�ṹ����
    Set mrsMainInfo = New ADODB.Recordset
    With mrsMainInfo
        .Fields.Append "���", adInteger, , adFldKeyColumn              '��������ʶ��Ϣ
        .Fields.Append "Index", adInteger, , adFldIsNullable                'Ϊ��ʱ��ʾ���ǿؼ�����
        .Fields.Append "��Ϣֵ", adVarChar, 2000, adFldIsNullable  '��Ϣ�ڼ���ʱ��ֵ
        .Fields.Append "ErrInfo", adVarChar, 4000, adFldIsNullable  '��Ϣ���Ϸ���ʾ��Ϣ��
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Private Sub cmdUp_Click()
    Dim lngTmp As Long
    Dim strTmp As Variant
    Dim strFile As String
    Dim strMainInfo As String
    
    If mlng��� <= mlngFirst Then
        MsgBox "���Ѿ��ǵ�һ��ͼƬ�ˣ�", vbInformation, gstrSysName
        Exit Sub
    Else
        mlng��� = mlng��� - 1
        Do While Not mlng��� < mlngFirst
            mrsMainInfo.Filter = "Index=" & mlng���
            If mrsMainInfo.EOF Then
                Screen.MousePointer = 11
                Call ReadPatPricture(mlng֤��id & "," & mlng���, imgLoad, strFile)
                Screen.MousePointer = 0
                mstrNote = GetImageNote(mlng֤��id, mlng���)
                If imgLoad.Picture <> 0 Then
                    lngTmp = mrsMainInfo.RecordCount + 1
                    If imgLoad.Picture <> 0 Then
                        strMainInfo = strFile & "|" & mstrNote
                        mrsMainInfo.AddNew Array("���", "Index", "��Ϣֵ"), Array(lngTmp, mlng���, strMainInfo)
                    End If
                End If
            End If
            If Not mrsMainInfo.EOF Then
                mrsMainInfo.Filter = "Index=" & mlng���
                strTmp = Split(mrsMainInfo!��Ϣֵ, "|")
                txtNote = strTmp(1)
                imgPicture.Picture = LoadPicture(strTmp(0))
                Exit Sub
            End If
            mlng��� = mlng��� - 1
        Loop
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrH  As Long
    Dim strMainInfo As String
    Dim strFile As String
    Dim lngTmp As Long
    Dim strTmp As Variant
    
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
    If mlngTop + Me.Height > lngScrH Then
        Me.Top = mlngTop - Me.Height
    Else
        Me.Top = mlngHeight + 2500
    End If
    Me.Left = mlngLeft
    Call InitBaseInfo
    If mintType = 0 Or mintType = 1 Then
        Screen.MousePointer = 11
        Call ReadPatPricture(mlng֤��id & "," & mlng���, imgLoad, strFile)
        Screen.MousePointer = 0
        mstrNote = GetImageNote(mlng֤��id, mlng���)
        mrsMainInfo.Filter = "Index=" & mlng���
        If mrsMainInfo.EOF Then
            lngTmp = mrsMainInfo.RecordCount + 1
            If imgLoad.Picture <> 0 Then
                strMainInfo = strFile & "|" & mstrNote
                mrsMainInfo.AddNew Array("���", "Index", "��Ϣֵ"), Array(lngTmp, mlng���, strMainInfo)
            End If
        End If
        mrsMainInfo.Filter = "Index=" & mlng���
        If Not mrsMainInfo.EOF Then
            strTmp = Split(mrsMainInfo!��Ϣֵ, "|")
            txtNote = strTmp(1)
            imgPicture.Picture = LoadPicture(strTmp(0))
        Else
            MsgBox "�ò���û��ͼƬ��Ϣ��", vbInformation, gstrSysName
        End If
    Else
        imgPicture.Picture = LoadPicture(mstrFile)
    End If
    If mintType <> 0 Then
        txtNote.Visible = False
        lblNote.Visible = False
        cmdUp.Visible = False
        cmdDown.Visible = False
        Me.Height = imgPicture.Height + imgPicture.Top + 600
    End If
End Sub

Public Function ShowMe(frmParent As Object, ByVal lng֤��ID As Long, ByVal intTYPE As Integer, ByVal X As Long, ByVal Y As Long, ByVal lngHeight As Long, Optional lng��� As Long, Optional strFile As String)
    Dim rsTmp  As New ADODB.Recordset
    Dim i As Long
    
    Set mfrmParent = frmParent
    mlng֤��id = lng֤��ID
    mlngLeft = X
    mlngTop = Y
    mlngHeight = lngHeight
    mintType = intTYPE
    mstrFile = strFile
    mlngFirst = lng���
    mlng��� = lng���
    If mintType = 0 Or mintType = 1 Then
        Set rsTmp = GetCertPicture(mlng֤��id, mlng���, 1)
        If rsTmp.EOF Then
            MsgBox "�ò���û��ͼƬ��Ϣ��", vbInformation, gstrSysName
        Else
            For i = 0 To rsTmp.RecordCount - 1
                If i = rsTmp.RecordCount - 1 Then
                    mlngCount = rsTmp!���
                End If
                rsTmp.MoveNext
            Next
            Me.Show 1, mfrmParent
        End If
    Else
        Me.Show 1, mfrmParent
    End If
End Function

Private Sub Form_Resize()
    imgPicture.Left = (Me.ScaleWidth / 2) + Me.ScaleLeft - (imgPicture.Width / 2) + Me.ScaleLeft
    txtNote.Left = imgPicture.Left
    lblNote.Left = txtNote.Left - lblNote.Width - 20
    cmdUp.Left = imgPicture.Left
    cmdDown.Left = imgPicture.Left + imgPicture.Width - cmdDown.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim objFile As New FileSystemObject
    
    mlng֤��id = 0
    mlng��� = 0
    mstrNote = ""
    mrsMainInfo.Filter = 0
    Do While Not mrsMainInfo.EOF
        If objFile.FileExists(Mid(mrsMainInfo!��Ϣֵ, 1, InStr(mrsMainInfo!��Ϣֵ, "|") - 1)) Then
            Kill Mid(mrsMainInfo!��Ϣֵ, 1, InStr(mrsMainInfo!��Ϣֵ, "|") - 1)
        End If
        mrsMainInfo.MoveNext
    Loop
    Set mrsMainInfo = Nothing
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If zlCommFun.ActualLen(txtNote.Text) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
        KeyAscii = 0
    End If
End Sub


