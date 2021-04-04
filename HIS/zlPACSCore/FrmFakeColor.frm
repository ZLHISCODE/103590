VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form FrmFakeColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "α������"
   ClientHeight    =   6870
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9870
   DrawStyle       =   1  'Dash
   Icon            =   "FrmFakeColor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      Caption         =   "�޸�α��Ԥ�跽��"
      Height          =   6015
      Left            =   4200
      TabIndex        =   25
      Top             =   120
      Width           =   5655
      Begin VB.ListBox lstColorList 
         Height          =   3120
         Left            =   3600
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "��ֹ"
         Height          =   975
         Left            =   3000
         TabIndex        =   28
         Top             =   4320
         Width           =   2500
         Begin VB.TextBox txtColor 
            Height          =   285
            Index           =   2
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   1300
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "��"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   29
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "λ�ã�"
            Height          =   180
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��ɫ��"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   540
         End
         Begin VB.Shape shpColor 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   2
            Left            =   960
            Top             =   600
            Width           =   1100
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "��ʼ"
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   4320
         Width           =   2500
         Begin VB.CommandButton cmdColor 
            Caption         =   "��"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   27
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtColor 
            Height          =   285
            Index           =   1
            Left            =   960
            TabIndex        =   12
            Top             =   240
            Width           =   1300
         End
         Begin VB.Shape shpColor 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   1
            Left            =   960
            Top             =   600
            Width           =   1100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��ɫ��"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "λ�ã�"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdModifyColor 
         Caption         =   "�޸�"
         Height          =   350
         Left            =   2400
         TabIndex        =   17
         Top             =   5520
         Width           =   1100
      End
      Begin VB.PictureBox picColorRect 
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3315
         ScaleWidth      =   3315
         TabIndex        =   9
         Top             =   480
         Width           =   3375
         Begin MSComDlg.CommonDialog dlgColor 
            Left            =   2160
            Top             =   2640
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ԥ��α�ʷ�����"
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton cmdDelPalette 
         Cancel          =   -1  'True
         Caption         =   "ɾ��"
         Height          =   350
         Left            =   2760
         TabIndex        =   3
         Top             =   840
         Width           =   1100
      End
      Begin VB.CommandButton cmdModifyPalette 
         Caption         =   "�޸�"
         Height          =   350
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddPalette 
         Caption         =   "����"
         Height          =   350
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   1100
      End
      Begin VB.ComboBox cobColor 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3672
      End
   End
   Begin VB.PictureBox picColorScale 
      Height          =   3345
      Left            =   3840
      ScaleHeight     =   3285
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ����(&D)"
      Height          =   350
      Left            =   3240
      TabIndex        =   21
      Top             =   6360
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ����(&U)"
      Height          =   350
      Left            =   1320
      TabIndex        =   20
      Top             =   6360
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   3690
      TabIndex        =   23
      Top             =   1680
      Width           =   3756
      Begin DicomObjects.DicomViewer ViewerFackColor 
         Height          =   3240
         Left            =   30
         TabIndex        =   4
         Top             =   45
         Width           =   3630
         _Version        =   262147
         _ExtentX        =   6403
         _ExtentY        =   5715
         _StockProps     =   35
         BackColor       =   -2147483640
      End
   End
   Begin VB.Frame famOpt 
      Caption         =   "Ӧ�÷�Χ(������)"
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   4005
      Begin VB.OptionButton OptImage 
         Caption         =   "����ͼ��"
         Height          =   288
         Index           =   2
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1116
      End
      Begin VB.OptionButton OptImage 
         Caption         =   "��ѡͼ��"
         Height          =   276
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1128
      End
      Begin VB.OptionButton OptImage 
         Caption         =   "��ǰͼ��"
         Height          =   240
         Index           =   0
         Left            =   228
         TabIndex        =   6
         Top             =   330
         Value           =   -1  'True
         Width           =   1164
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7080
      TabIndex        =   19
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5160
      TabIndex        =   18
      Top             =   6360
      Width           =   1100
   End
End
Attribute VB_Name = "FrmFakeColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public f As frmViewer
Dim varColor As Variant
Dim varTmpColor As Variant
Dim intX As Integer
Dim intY As Integer

Private Sub cmdAddPalette_Click()
    Dim strPaletteName As String
    Dim dsData As New Recordset
    Dim strSQL As String
    Dim varNewPalette As String
    Dim intID As Integer
    On Error GoTo errh
    '���α��ģ�������Ƿ����
    strPaletteName = InputBox("�������µ�α��ģ�����ƣ�", "����α��ģ��")
    If strPaletteName = "" Then Exit Sub
    '���α��ģ�������Ƿ��ظ�
    If blLocalRun = True Then
        strSQL = "select ��ɫ from Ӱ����ɫ�嵥 "
        Set dsData = cnAccess.Execute(strSQL)
    Else
        strSQL = "select ��ɫ from Ӱ����ɫ�嵥 "
        Set dsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    While Not dsData.EOF
        If dsData!��ɫ = strPaletteName Then
            MsgBox "α��ģ�������ظ������������롣", vbInformation, gstrSysName
            Exit Sub
        End If
        dsData.MoveNext
    Wend
    dsData.Close
    '����lstBox�����ݣ���֯��ɫ���ݿ�����
    varNewPalette = funGetPalette()
    If blLocalRun = True Then
        '���µ�α��ģ�汣�浽���ݿ�
        strSQL = "insert into Ӱ����ɫ�嵥 (��ɫ,��ɫ����) values ('" & strPaletteName & "','" & varNewPalette & "')"
        cnAccess.Execute strSQL
        strSQL = "select ���,��ɫ���� from Ӱ����ɫ�嵥 where ��ɫ = '" & strPaletteName & "'"
        dsData.Open strSQL, cnAccess, adOpenDynamic, adLockPessimistic
        '��α��ģ��������ӵ���ɫ�������б���
        Me.cobColor.AddItem "�û�������" & strPaletteName
        Me.cobColor.ItemData(Me.cobColor.NewIndex) = dsData!���
    Else
        strSQL = "select max(���) as ������ from Ӱ����ɫ�嵥 "
        Set dsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        intID = dsData("������") + 1
        strSQL = "ZL_Ӱ����ɫ�嵥_INSERT('" & strPaletteName & "','" & varNewPalette & "',0)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Me.cobColor.AddItem "�û�������" & strPaletteName
        Me.cobColor.ItemData(Me.cobColor.NewIndex) = intID
    End If
    Exit Sub
errh:
    If blLocalRun = False Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    Else
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    End If
End Sub

Private Function funGetPalette() As String
'------------------------------------------------
'���ܣ� ��lstColorList�ж�ȡ��ǰ���úõĵ�ɫ������
'������ ��
'���أ� ��ɫ����ɫ��
'------------------------------------------------
    Dim i As Integer
    Dim bytR As Byte
    Dim bytG As Byte
    Dim bytB As Byte
    Dim strColor As String
    Dim strRGB As String
    
    For i = 1 To 256
        strColor = Me.lstColorList.list(i - 1)
        strColor = Right(strColor, Len(strColor) - InStr(strColor, "|") - 2)
        bytR = strColor Mod 256
        bytG = strColor \ 256 Mod 256
        bytB = strColor \ 256 \ 256
        
        funGetPalette = funGetPalette & bytR & "," & bytG & "," & bytB & ";"
    Next i
End Function
Private Sub cmdColor_Click(Index As Integer)
    Me.dlgColor.Color = Me.shpColor(Index).FillColor
    Me.dlgColor.ShowColor
    Me.shpColor(Index).FillColor = Me.dlgColor.Color
End Sub

Private Sub cmdDelPalette_Click()
    Dim strSQL As String
    Dim rsData As Recordset
    If Me.cobColor.ListIndex = -1 Then Exit Sub
    On Error GoTo errh
    If MsgBox("�Ƿ�ȷ��ɾ��ģ�棺" & Me.cobColor.list(Me.cobColor.ListIndex), vbQuestion + vbOKCancel, gstrSysName) = vbOK Then
        '�ж�ģ���Ƿ�ϵͳģ�棬ϵͳģ�治����ɾ��
        If blLocalRun = True Then
            strSQL = "select ϵͳ���� from Ӱ����ɫ�嵥 where ��� =" & Me.cobColor.ItemData(Me.cobColor.ListIndex)
            Set rsData = cnAccess.Execute(strSQL)
        Else
            strSQL = "select ϵͳ���� from Ӱ����ɫ�嵥 where ��� = [1]"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Me.cobColor.ItemData(Me.cobColor.ListIndex)))
        End If
        
        If Not rsData.BOF And Not rsData.EOF Then
            If rsData!ϵͳ���� = 1 Then
                MsgBox "��ǰѡ��ģ��Ϊϵͳģ�棬������ɾ����ֻ��ɾ���û��Լ�������ģ�档", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '�����ݿ���ɾ��ģ��
        If blLocalRun = True Then
            strSQL = "delete from Ӱ����ɫ�嵥 where ���=" & Me.cobColor.ItemData(Me.cobColor.ListIndex)
            cnAccess.Execute strSQL
        Else
            strSQL = "ZL_Ӱ����ɫ�嵥_DELETE(" & Me.cobColor.ItemData(Me.cobColor.ListIndex) & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        'ɾ��ģ�������б��е�ģ������
        Me.cobColor.RemoveItem Me.cobColor.ListIndex
        '���������б�ĵ�ǰֵ
        Me.cobColor.ListIndex = 0
    End If
    Exit Sub
errh:
    If blLocalRun = False Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    Else
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    End If
End Sub

Private Sub cmdModifyColor_Click()
    If Me.txtColor(1).Text = "" Or Me.txtColor(2).Text = "" Then
        MsgBox "��������ɫ�Ŀ�ʼ�ͽ���ֵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Dim i As Integer
    Dim bytR As Byte
    Dim bytG As Byte
    Dim bytB As Byte
    Dim bytStartR As Integer
    Dim bytStartG As Integer
    Dim bytStartB As Integer
    Dim bytEndR As Integer
    Dim bytEndG As Integer
    Dim bytEndB As Integer
    Dim lngColor As Long
    Dim intFlagLength As Integer        '��¼��ʼ����ֹ��ɫ����ɫ���֮��ľ���
    
    intFlagLength = Me.txtColor(2).Text - Me.txtColor(1).Text
    If intFlagLength = 0 Then intFlagLength = 1
    bytStartR = Me.shpColor(1).FillColor Mod 256
    bytStartG = Me.shpColor(1).FillColor \ 256 Mod 256
    bytStartB = Me.shpColor(1).FillColor \ 256 \ 256
    bytEndR = Me.shpColor(2).FillColor Mod 256
    bytEndG = Me.shpColor(2).FillColor \ 256 Mod 256
    bytEndB = Me.shpColor(2).FillColor \ 256 \ 256
    For i = Me.txtColor(1).Text To Me.txtColor(2).Text
        bytR = bytStartR + (bytEndR - bytStartR) * ((i - Me.txtColor(1).Text) / intFlagLength)
        bytG = bytStartG + (bytEndG - bytStartG) * (i - Me.txtColor(1).Text) / intFlagLength
        bytB = bytStartB + (bytEndB - bytStartB) * (i - Me.txtColor(1).Text) / intFlagLength
        varTmpColor(i) = bytR & "," & bytG & "," & bytB
    Next i
    picColorRect_Paint
    subFillColorList
    
End Sub

Private Sub cmdModifyPalette_Click()
    If Me.cobColor.ListIndex = -1 Then Exit Sub
    Dim strSQL As String
    Dim rsData As New Recordset
    Dim varNewPalette As String
    On Error GoTo errh
    If blLocalRun = True Then
        '��鵱ǰ��ѡ�еĵ�ɫ���Ƿ�Ϊϵͳģ�棬����ǣ�����ʾ�������޸�ϵͳģ��
        strSQL = "select ϵͳ���� from Ӱ����ɫ�嵥 where ���=" & Me.cobColor.ItemData(Me.cobColor.ListIndex)
        Set rsData = cnAccess.Execute(strSQL)
    Else
        strSQL = "select ϵͳ���� from Ӱ����ɫ�嵥 where ���= [1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Me.cobColor.ItemData(Me.cobColor.ListIndex)))
    End If
    If Not rsData.BOF And Not rsData.EOF Then
        If rsData!ϵͳ���� = 1 Then
            MsgBox "��ǰѡ��ģ��Ϊϵͳģ�棬�������޸ģ�ֻ���޸��û��Լ�������ģ�档", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    rsData.Close
    '�޸��û���ɫ��ģ��
    varNewPalette = funGetPalette
    
    '���µ�α��ģ�汣�浽���ݿ�
    If blLocalRun = True Then
        strSQL = "update Ӱ����ɫ�嵥 set ��ɫ���� = '" & varNewPalette & "' where ��� = " & Me.cobColor.ItemData(Me.cobColor.ListIndex)
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_Ӱ����ɫ�嵥_UPDATE(" & Me.cobColor.ItemData(Me.cobColor.ListIndex) & ",'" & varNewPalette & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    cobColor_Click
    Exit Sub
    
errh:
    If blLocalRun = False Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    Else
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    End If
End Sub

Private Sub cmdNext_Click()
    If cobColor.ListIndex < cobColor.ListCount - 1 Then cobColor.ListIndex = cobColor.ListIndex + 1
End Sub

Private Sub cmdOK_Click()
    Dim iColorNum As Long
    Dim strTemp As String
    Dim tmpF As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim NewImgs As New DicomImages
    Dim im As DicomImage, imb As New DicomImage
    Dim blnSelected As Boolean      '�Ƿ�����ѡ���ͼ��
    Dim NewImg As DicomImage
    Dim i As Integer
    Dim iVieweIndex As Integer
    Dim iImageIndex As Integer
    
    iVieweIndex = f.intSelectedSerial
    '��Ҫ����ѡ���������ԡ���ǰͼ�񡱣�����ѡͼ�񡱣�������ͼ����α��
    '���ѡ���˶���ѡͼ�����α�ʴ�������Ƿ����Ѿ���ѡ���ͼ��
    If Me.OptImage(1) Then
        For i = 1 To ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count
            If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnSelected = True Then
                blnSelected = True
                Exit For
            End If
        Next i
        
        If blnSelected = False Then
            MsgBox "��ǰû��ѡ���κ�ͼ�񣬲�����������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�ڹ�Ƭվ����ʾͼ��
    '���ж�ͼ���ѡ�����������ǡ���ѡͼ�񡱺͡�����ͼ����Ҫ���Ȱ���Щͼ����ؽ���
    If OptImage(1).Value = True Then    '��ѡͼ��
        iImageIndex = 1
        For i = 1 To ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count
            If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnSelected = True Then
                '�����ж�ͼ���Ƿ��Ѿ�װ�أ�����Ѿ�װ�أ����ҵ����ͼ����ʾ���������û��װ�أ���װ�ظ�ͼ��
                If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnDisplayed = False Then
                    Call funcAddAImageA(f.Viewer(iVieweIndex), i)
                End If
                
                '����ͼ�������
                While f.Viewer(iVieweIndex).Images(iImageIndex).Tag < i And iImageIndex < f.Viewer(iVieweIndex).Images.Count
                    iImageIndex = iImageIndex + 1
                Wend
                
                If iImageIndex <= f.Viewer(iVieweIndex).Images.Count Then
                    If f.Viewer(iVieweIndex).Images(iImageIndex).Tag = i Then
                        Set im = f.Viewer(iVieweIndex).Images(iImageIndex)
                    End If
                End If
                
                If Not im Is Nothing Then
                    imgs.Add im
                End If
            End If
        Next i
    ElseIf OptImage(2).Value = True Then    '����ͼ��
        'ȷ������ͼ�󶼱����ص�Viewer��
        If ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count <> f.Viewer(iVieweIndex).Images.Count Then
            For i = 1 To ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count
                If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnDisplayed = False Then
                    Call funcAddAImageA(f.Viewer(iVieweIndex), i)
                End If
            Next i
        End If
        '��Viewer�е�����ͼ�󶼼��ص�imgs������
        For i = 1 To f.Viewer(iVieweIndex).Images.Count
            imgs.Add f.Viewer(iVieweIndex).Images(i)
        Next i
    Else    '��ǰͼ��
        imgs.Add f.SelectedImage
    End If
    
    '����ͼ�񼯺��е�����ͼ���޸ĳ�α��ͼ��
    iColorNum = cobColor.ItemData(cobColor.ListIndex)
    Call GetBmpPaletteFromDB(iColorNum)
    For Each im In imgs
        im.Labels.Clear
        strTemp = App.Path & "\temp\" & tmpF.GetTempName
        im.FileExport strTemp, "BMP"
        If ChangeBmpPaletteFromDB(strTemp) = 1 Then MsgBox "��ͼ���Ѿ��ǲ�ɫͼ�񣬲��ܹ�����α�ʲ�����", vbInformation, gstrSysName
        Set NewImg = New DicomImage
        NewImg.FileImport strTemp, "BMP"
        NewImgs.Add NewImg
        
        'װ����ϣ���������ʱͼ��ͬʱɾ��
        Kill strTemp
        DoEvents
        Me.Caption = "α�����á������������ڴ���α�ʣ����Ժ�"
    Next
    
    '���������ʾ��Viewer��������2�������ú�����ʾ2��
    If f.intCountX < 2 Then
        f.intCountX = 2
        Call subChangeSeriesLayout(f)
    End If
    
    '������ʾNewImgs�е�ͼ��
    Call funShowTempImages(f, NewImgs, 0)
    
    Unload Me
End Sub

Private Sub cmdPrevious_Click()
    If cobColor.ListIndex > 0 Then cobColor.ListIndex = cobColor.ListIndex - 1
End Sub
Private Sub cobColor_Click()
    Dim tmpF As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim iPalNum As Long
    Dim strTemp As String
    Dim DirTemp As SECURITY_ATTRIBUTES              '����Ŀ¼ʱ��Ҫ������
    Dim CreateTrue As Integer                        '����Ŀ¼�Ƿ�ɹ�����0��ʾ�ɹ���

    imgs.Add f.SelectedImage
    imgs(1).Labels.Clear
    iPalNum = cobColor.ItemData(cobColor.ListIndex)
    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        '����Ŀ¼
        CreateTrue = CreateDirectory(App.Path & "\temp", DirTemp)
        If CreateTrue = 0 Then
            '����Ŀ¼ʧ��ʱ�˳�
            MsgBox "������ʱĿ¼" & App.Path & "\temp" & "ʧ��!", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    Call GetBmpPaletteFromDB(iPalNum)
    
    strTemp = App.Path & "\temp\" & tmpF.GetTempName
    imgs(1).FileExport strTemp, "BMP"
    If ChangeBmpPaletteFromDB(strTemp) Then MsgBox "��ͼ���Ѿ��ǲ�ɫͼ�񣬲��ܹ�����α�ʲ�����", vbInformation, gstrSysName
    Dim img As New DicomImage
    img.FileImport strTemp, "BMP"
    ViewerFackColor.Images.Clear
    ViewerFackColor.Images.Add img
    ViewerFackColor.Refresh
End Sub

Private Function GetBmpPaletteFromDB(palNum As Long) As Integer
'------------------------------------------------
'���ܣ�����α����ɫ��ţ������ݿ��ж�ȡα����ɫ���������ػ���ɫ��ʾ�б�
'������palNum α�ʷ����е�α����ɫ���
'���أ�0-�޸���ImgFile�ļ��ĵ�ɫ��ɹ���1������ͼ��Ϊ��ɫͼ�񣬲����޸ĵ�ɫ�塣
'------------------------------------------------
    Dim vTemp As Variant
    Dim strSQL As String
    Dim rsFackColor As Recordset
    Dim strFackColor As String
    
    On Error GoTo 0
    
    If blLocalRun = True Then
        strSQL = "select ��ɫ���� from Ӱ����ɫ�嵥 where ���=" & palNum
        Set rsFackColor = cnAccess.Execute(strSQL)
    Else
        strSQL = "select ��ɫ���� from Ӱ����ɫ�嵥 where ���=[1]"
        Set rsFackColor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, palNum)
    End If
    strFackColor = rsFackColor.Fields("��ɫ����")
    vTemp = Split(strFackColor, ";")
    varColor = vTemp
    varTmpColor = varColor
    
    picColorScale_Paint 'ˢ��α�ʱ�ߵ���ʾ
    picColorRect_Paint  'ˢ��α����ɫ���ε���ʾ
    subFillColorList    'ˢ�²������ɫ�б����ʾ

End Function

Private Function ChangeBmpPaletteFromDB(ImgFile As String) As Integer
'------------------------------------------------
'���ܣ��޸�8λBMPͼ��ĵ�ɫ�壬�Ӷ�ʵ��α�ʵĹ���
'������ImgFile ��Ҫʵ��α�ʵ�8λ�Ҷ�bmpͼ���ļ���
'���أ�0-�޸���ImgFile�ļ��ĵ�ɫ��ɹ���1������ͼ��Ϊ��ɫͼ�񣬲����޸ĵ�ɫ�塣
'�ϼ���������̣�FrmFakeColor.cmdOK_Click��FrmFakeColor.cobColor_Click
'�¼���������̣���
'�����ˣ��ƽ�
'˵���� ����ʹ�õ�α�ʵ�ɫ��ֱ�Ӹ���ɫ������һ�𱣴������ݿ�ZLPACS.MDB�С���ɫ�嵥�������棬
'       ���й�����51��α�ʵ�ɫ�塣ʹ��WINDOWS GDI API��LOGPALETTE�ṹ�ĸ�ʽ�洢�����滹���ĸ�0��
'       ÿһ����ɫ��ʵ�ʰ������ֽ���Ϊ1032���ֽڡ�
'------------------------------------------------
    Dim palR As Byte
    Dim palG As Byte
    Dim palB As Byte
    Dim palFlag As Byte
    Dim i As Integer
    Dim intImageType As Integer
    Dim intRGB As String

    ChangeBmpPaletteFromDB = 0
    On Error Resume Next
    '��ͼ���ļ�
    Open ImgFile For Binary As #1
    
    '�ж�ͼ���Ƿ�ڰ�ͼ��
    Get #1, 29, intImageType
    If intImageType > 8 Then
        ChangeBmpPaletteFromDB = 1
        Close #1
        Exit Function
    End If
    
    On Error Resume Next
    
    For i = 0 To 255
        intRGB = varColor(i)
        '�ӵ�ɫ���ļ��ж�ȡR,G,B�ͱ�־λ������λ�����һ����ɫ������ɫ���ܹ�ʹ����256����ɫ
        palR = strGetRGB(intRGB, 0)
        palG = strGetRGB(intRGB, 1)
        palB = strGetRGB(intRGB, 2)
        palFlag = strGetRGB(intRGB, 3)
        
        '��α�ʵ�ɫ��д�뵽ͼ���ļ���Ӧλ���У�BMP�ļ�ǰ14λ��BITMAPFILEHEADER,
        '���40λ��BITMAPINFO��BITMAPINFO =BITMAPINFOHEADER + RGBQUAD
        '��54λ��ʼ������ͼ��ĵ�ɫ�壬������ǽ�ͼ��ԭ�еĵ�ɫ�廻��α�ʵ�ɫ��
        Put #1, 54 + 4 * i + 1, palB
        Put #1, 54 + 4 * i + 2, palG
        Put #1, 54 + 4 * i + 3, palR
        Put #1, 54 + 4 * i + 4, palFlag
    Next
    Close #1
    Exit Function
err:
    Close #1
End Function

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.cobColor.ListIndex = -1
    Me.picColorScale.Line -(10, 10)
    If InStr(mstrPrivs, "��Ƭ����") <> 0 Then
        cmdAddPalette.Enabled = True
        cmdModifyPalette.Enabled = True
        cmdDelPalette.Enabled = True
        Frame2.Enabled = True
    Else
        cmdAddPalette.Enabled = False
        cmdModifyPalette.Enabled = False
        cmdDelPalette.Enabled = False
        Frame2.Enabled = False
    End If
End Sub

Private Sub lstColorList_Click()
    Dim strColor As String
    strColor = Me.lstColorList.list(Me.lstColorList.ListIndex)
    Me.txtColor(1).Text = Val(left(strColor, InStr(strColor, "|") - 3))
    Me.txtColor(2).Text = Me.txtColor(1).Text
    strColor = Right(strColor, Len(strColor) - InStr(strColor, "|") - 2)
    Me.shpColor(1).FillColor = Val(strColor)
    Me.shpColor(2).FillColor = Val(strColor)
End Sub

Private Sub picColorRect_Click()
    Dim intXColor As Integer
    Dim intYColor As Integer
    Dim intColor As Integer
    Dim strColor As String
    intXColor = intX \ (Me.picColorRect.width / 16)
    intYColor = intY \ (Me.picColorRect.height / 16)
    If intYColor >= 16 Then intYColor = 15
    If intXColor >= 16 Then intXColor = 15
    intColor = intYColor * 16 + intXColor
    
    Me.txtColor(1).Text = intColor
    Me.txtColor(2).Text = intColor
    
    strColor = Me.lstColorList.list(intColor)
    strColor = Right(strColor, Len(strColor) - InStr(strColor, "|") - 2)
    Me.shpColor(1).FillColor = Val(strColor)
    Me.shpColor(2).FillColor = Val(strColor)
End Sub

Private Sub picColorRect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    intX = x
    intY = y
End Sub

Private Sub picColorRect_Paint()
    Dim i As Integer
    Dim j As Integer
    Dim intColor As Integer
    Dim intOldScaleWidth As Integer
    Dim intOldScaleHeight As Integer
    Dim intOldx As Integer
    Dim intOldy As Integer
    Dim intRGB As String
    
    '����ԭ�е�����
    intOldScaleWidth = Me.picColorRect.ScaleWidth
    intOldScaleHeight = Me.picColorRect.ScaleHeight
    intOldx = Me.picColorRect.CurrentX
    intOldy = Me.picColorRect.CurrentY
    '���������ó�256�У�10��
    Me.picColorRect.Scale (0, 0)-(160, 160)
    
    '����α�ʱ�ߵĵ�λ�߶�
    Me.picColorRect.DrawWidth = 1
    Me.picColorRect.FillStyle = 0
    Me.picColorRect.ForeColor = vbBlack
    intColor = 0
    
    For j = 0 To 15
        For i = 0 To 15
            intRGB = varTmpColor(intColor)
            Me.picColorRect.FillColor = RGB(strGetRGB(intRGB, 0), strGetRGB(intRGB, 1), strGetRGB(intRGB, 2))
            Me.picColorRect.Line (i * 10 + 1, j * 10 + 1)-(i * 10 + 9, j * 10 + 9), , B
            intColor = intColor + 1
        Next i
    Next j
    
    '������ԭ
    Me.picColorRect.Scale (0, 0)-(intOldScaleWidth, intOldScaleHeight)
    Me.picColorRect.CurrentX = intOldx
    Me.picColorRect.CurrentY = intOldy
End Sub

Private Sub picColorScale_Paint()
    Dim i As Integer
    Dim intOldScaleWidth As Integer
    Dim intOldScaleHeight As Integer
    Dim intOldx As Integer
    Dim intOldy As Integer
    Dim intRGB As String
    
    '����ԭ�е�����
    intOldScaleWidth = Me.picColorScale.ScaleWidth
    intOldScaleHeight = Me.picColorScale.ScaleHeight
    intOldx = Me.picColorScale.CurrentX
    intOldy = Me.picColorScale.CurrentY
    '���������ó�256�У�10��
    Me.picColorScale.Scale (0, 0)-(9, 255)
    
    '����α�ʱ�ߵĵ�λ�߶�
    Me.picColorScale.DrawWidth = 1
    
    For i = 0 To 255
        intRGB = varColor(i)
        Me.picColorScale.ForeColor = RGB(strGetRGB(intRGB, 0), strGetRGB(intRGB, 1), strGetRGB(intRGB, 2))
        Me.picColorScale.Line (0, i)-(10, i)
    Next i
    
    '������ԭ
    Me.picColorScale.Scale (0, 0)-(intOldScaleWidth, intOldScaleHeight)
    Me.picColorScale.CurrentX = intOldx
    Me.picColorScale.CurrentY = intOldy
End Sub

Private Sub subFillColorList()
    Dim i As Integer
    Dim intRGB As String
    Me.lstColorList.Clear  '���listBox��ԭ������
    For i = 0 To 255
        intRGB = varTmpColor(i)
        Me.lstColorList.AddItem Val(i) & "  |  " & RGB(strGetRGB(intRGB, 0), strGetRGB(intRGB, 1), strGetRGB(intRGB, 2))
    Next i
End Sub

Private Sub txtColor_GotFocus(Index As Integer)
    Me.txtColor(Index).SelStart = 0
    Me.txtColor(Index).SelLength = Len(Me.txtColor(Index).Text)
End Sub

Private Sub txtColor_LostFocus(Index As Integer)
    If Val(Me.txtColor(Index).Text) = 0 Then    'ֻ������������
        Me.txtColor(Index).Text = 0
    End If
    'ȷ����ֵ��0-255֮��
    If Val(Me.txtColor(Index).Text) < 0 Then Me.txtColor(Index).Text = 0
    If Val(Me.txtColor(Index).Text) > 255 Then Me.txtColor(Index).Text = 255
    'ȷ������txt֮����ֵ�Ĵ�С��ϵ
    If Index = 1 Then
        If Val(Me.txtColor(1).Text) > Val(Me.txtColor(2).Text) Then Me.txtColor(1).Text = Me.txtColor(2).Text
    Else
        If Val(Me.txtColor(2).Text) < Val(Me.txtColor(1).Text) Then Me.txtColor(2).Text = Me.txtColor(1).Text
    End If
End Sub
Private Function strGetRGB(strRGB As String, intRGB As Integer) As Integer
    '���ִ��еõ�RGB��ɫ
    Dim StrTmp As Variant
    StrTmp = Split(strRGB, ",")
    strGetRGB = StrTmp(intRGB)
End Function

