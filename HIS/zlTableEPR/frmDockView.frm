VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Begin VB.Form frmDockView 
   BorderStyle     =   0  'None
   Caption         =   "Ԥ������"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicDy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   0
      Left            =   1140
      ScaleHeight     =   885
      ScaleWidth      =   1170
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   1170
   End
   Begin TTF160Ctl.F1Book F1Main 
      Height          =   1305
      Left            =   180
      TabIndex        =   0
      Top             =   105
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2302
      _0              =   $"frmDockView.frx":0000
      _1              =   $"frmDockView.frx":0409
      _2              =   $"frmDockView.frx":0812
      _3              =   $"frmDockView.frx":0C1B
      _4              =   $"frmDockView.frx":1024
      _count          =   5
      _ver            =   2
   End
End
Attribute VB_Name = "frmDockView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Doc As cTableEPR, mblnInit As Boolean
Private Sub F1Main_TopLeftChanged()
Dim i As Integer, strCellKey As String
    On Error GoTo errHand
    If mblnInit Then Exit Sub
    For i = 1 To PicDy.UBound
        If ChkControl(PicDy(i)) Then
            If PicDy(i).Picture.Handle <> 0 Then
                strCellKey = Split(PicDy(i).Tag, "|")(1)
                Call PaintPictureOnTable(strCellKey)
            End If
        End If
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    F1Main.Top = 0: F1Main.Left = 0: F1Main.Width = Me.ScaleWidth: F1Main.Height = Me.ScaleHeight
End Sub
Public Sub zlRefresh(Tmp As cTableEPR)
'���ܣ�ˢ�½���
Dim l As Long, lCount As Long
    On Error GoTo errHand
    '�崰ͼƬ�ؼ�
    Set Doc = Tmp
    PicDy(0).Visible = False
    For l = 1 To PicDy.UBound
        If ChkControl(PicDy(l)) Then
            Unload PicDy(l)
        End If
    Next
    
    With F1Main '��ʼ�����
        .DeleteRange .MinRow, .MinCol, .MaxRow, .MaxCol, F1ShiftRows
        .MaxCol = 4: .MaxRow = 4
    End With
    
    If Doc.ReadFileStructure Then   '��ȡ�ļ��ṹ
        Doc.ReadFileContent Doc.mblnMove   '��ȡ�ļ�����
    Else
        Exit Sub
    End If
    Call RefreshF1Main
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub RefreshF1Main()
Dim lngRow As Long, lngCol As Long, lngCell As Long, vCell As F1CellFormat, lngCount As Long, strShow As String
    mblnInit = True
    With F1Main
        .DeleteRange .MinRow, .MinCol, .MaxRow, .MaxCol, F1ShiftRows
        .ShowTabs = F1TabsOff
        .AllowMoveRange = False '�ƶ�ѡ������
        .AllowFillRange = False '�϶���Χ��ֵ,���¼����ɿ���
        .AllowInCellEditing = False '��Ԫ��༭
        .AllowEditHeaders = False '�༭��ͷ
        .AllowDesigner = False  '�������
        .AllowDelete = False '��ʾ��Ӣ�ĵģ���ò�Ҫ���������ͨ��KeyDown����
        .ShowLockedCellsError = False '��������Ԫ����б༭ʱ����Ϣ��ʾ
        .ScrollToLastRC = False '������������һ����Ԫ��
        .ColWidthUnits = F1ColWidthUnitsTwips '�п���㵥λΪ��
        .DefaultFontName = "����"
        .DefaultFontSize = 9
        .MaxCol = Doc.Cells.Cols
        .MaxRow = Doc.Cells.Rows

        '���и��п�
        For lngRow = 1 To .MaxRow
            .RowHeight(lngRow) = Doc.Cells.Cell(lngRow, 1).Height
        Next
        For lngCol = 1 To .MaxCol
            .ColWidthTwips(lngCol) = Doc.Cells.Cell(1, lngCol).Width
        Next
        
        lngCount = Doc.Cells.Count
        For lngCell = 1 To lngCount
            lngRow = Doc.Cells(lngCell).Row: lngCol = Doc.Cells(lngCell).Col
            With Doc.Cells.Cell(lngRow, lngCol)
                'ָ������
                If .Merge And InStr(.MergeRange, ";") > 0 Then 'MergeRange���ݸ�ʽ (���Ϸ�)��,��;(���·�)��,��
                    F1Main.SetSelection Split(Split(.MergeRange, ";")(0), ",")(0), Split(Split(.MergeRange, ";")(0), ",")(1), Split(Split(.MergeRange, ";")(1), ",")(0), Split(Split(.MergeRange, ";")(1), ",")(1)
                Else
                    F1Main.SetSelection lngRow, lngCol, lngRow, lngCol
                End If
                Set vCell = F1Main.CreateNewCellFormat
                If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then 'ֻ�кϲ���Ԫ���׸���Ǻϲ���Ԫ���ˢ��
'                    vCell.ProtectionLocked = .��������  '�Ƿ�����,��������,�п�,�п�,ǩ��������ʱд��Database
                    vCell.MergeCells = .Merge
                    vCell.WordWrap = True
                    vCell.FontName = .FontName          '����>����</����>
                    vCell.FontSize = .FontSize          '<�ֺ�>9</�ֺ�>
                    vCell.FontBold = .FontBold          '<����>False</����>
                    vCell.FontItalic = .FontItalic        '<б��>False</б��>
                    vCell.FontUnderline = .FontUnderline     '<�»���>False</�»���>
                    vCell.FontStrikeout = .FontStrikeout    '<ɾ����>False</ɾ����>
                    vCell.FontColor = .FontColor         '<������ɫ>vbblack</������ɫ>
                    vCell.AlignHorizontal = .HAlignment       '<�������>F1HAlignCenter</�������>
                    vCell.AlignVertical = .VAlignment       '<�������>F1VAlignCenter</�������>

                    Select Case .��������
                        Case cprCTFixtext    '0-�̶��ı�(���ɱ༭)
                            F1Main.TextRC(lngRow, lngCol) = .�����ı�
                        Case cprCTText '1-�ı���(�ɱ༭�����ı�)
                            F1Main.TextRC(lngRow, lngCol) = .�����ı�
                        Case cprCTElement    '2-��Ҫ��
                            If Doc.ET = TabET_�����ļ����� Or Doc.ET = TabET_ȫ��ʾ���༭ Then
                                If .ElementKey <> "" Then
                                    If Doc.Elements("K" & .ElementKey).������̬ = 1 Then
                                        F1Main.TextRC(lngRow, lngCol) = Doc.Elements("K" & .ElementKey).�����ı�
                                    Else
                                        F1Main.TextRC(lngRow, lngCol) = "[" & Doc.Elements("K" & .ElementKey).Ҫ������ & "]" & Doc.Elements("K" & .ElementKey).Ҫ�ص�λ
                                    End If
                                End If
                            Else
                                strShow = ""
                                If .�����ı� = "" Then
                                    If Doc.Elements("K" & .ElementKey).�滻�� = 1 Then '�Զ��滻Ҫ��
                                        strShow = GetReplaceEleValue(Doc.Elements("K" & .ElementKey).Ҫ������, Doc.EPRPatiRecInfo.����ID, Doc.EPRPatiRecInfo.��ҳID, Doc.EPRPatiRecInfo.������Դ, Doc.EPRPatiRecInfo.ҽ��id)
                                        If strShow = "" And Not Doc.Elements("K" & .ElementKey).�Զ�ת�ı� Then 'ûȡ��ֵ���Ƿ��Զ�ת�����ı�(��)
                                            strShow = "[" & Doc.Elements("K" & .ElementKey).Ҫ������ & "]" & Doc.Elements("K" & .ElementKey).Ҫ�ص�λ
                                        Else
                                            Doc.Elements("K" & .ElementKey).�����ı� = strShow
                                            .�����ı� = strShow & Doc.Elements("K" & .ElementKey).Ҫ�ص�λ
                                            strShow = .�����ı�
                                        End If
                                    Else
                                        If Doc.Elements("K" & .ElementKey).������̬ = 1 And Doc.Elements("K" & .ElementKey).Ҫ������ <> 2 Then '������̬=չ��
                                            .�����ı� = Doc.Elements("K" & .ElementKey).�����ı� & Doc.Elements("K" & .ElementKey).Ҫ�ص�λ
                                            strShow = .�����ı�
                                        Else
                                            strShow = "[" & Doc.Elements("K" & .ElementKey).Ҫ������ & "]" & Doc.Elements("K" & .ElementKey).Ҫ�ص�λ
                                        End If
                                    End If
                                    F1Main.TextRC(lngRow, lngCol) = strShow
                                Else
                                    F1Main.TextRC(lngRow, lngCol) = .�����ı�
                                End If
                            End If
                        Case cprCTTextElement '3-�ı����Ҫ�ػ�ϱ༭
                            GetTextELement .Key     '����Text Element��дF1Main�еĵ�Ԫ����������ı�
                        Case cprCTReportPic, cprCTPicture    '5-����ͼ
                            If Doc.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                Call PaintPictureOnTable(.Key)
                            End If
                            F1Main.TextRC(lngRow, lngCol) = IIf(.�������� = cprCTPicture, "�ο�ͼ", "����ͼ")
                        Case cprCTSign         '6-ǩ��'ǩ�������ʱ��Ϊռλ,��ʵ����Ϣ��û��ǩ��ʱ��ֹ��=0����ͨǩ�������ʱ����ʾ���Ա��ٴ�ǩ�����п�/�п�ǩ�������ʱҪ��ʾ
                            strShow = ""
                            If Doc.ET = TabET_�������༭ Or Doc.ET = TabET_��������� Then
                                'mReadOnly 0-����,1-ǩ������޸�,2-������򿪲��Ļ��������ǩ���汾
                                If .��ֹ�� <> 0 Then
                                    With Doc.Signs("K" & .SignKey)
                                        strShow = .ǰ������ & .���� & IIf(.��ʾ��ǩ, "����ǩ��_____________", "")
                                        strShow = strShow & IIf(Trim(.��ʾʱ��) = "", "", "��" & Format(.ǩ��ʱ��, .��ʾʱ��))
                                    End With
                                Else
                                    strShow = "[ǩ��λ]"
                                End If
                            Else
                                strShow = "[ǩ��λ]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow 'ǰ������ & ���� & ��ʾ��ǩ & ��ʾʱ��<>""(format(ǩ��ʱ��,��ʾʱ��)
                        Case cprCTRowSign, cprCTColSign '7-�п�ǩ�� '8-�п�ǩ��
                            strShow = ""
                            If Doc.ET = TabET_�������༭ Or Doc.ET = TabET_��������� Then
                                If .��ֹ�� <> 0 Then
                                    With Doc.Signs("K" & .SignKey)
                                        strShow = .ǰ������ & .���� & IIf(.��ʾ��ǩ, "����ǩ��_____________", "")
                                        strShow = strShow & IIf(Trim(.��ʾʱ��) = "", "", "��" & Format(.ǩ��ʱ��, .��ʾʱ��))
                                    End With
                                Else
                                    strShow = "[ǩ��λ]"
                                End If
                            Else
                                strShow = "[ǩ��λ]"
                            End If
                            F1Main.TextRC(lngRow, lngCol) = strShow 'ǰ������ & ���� & ��ʾ��ǩ & ��ʾʱ��<>""(format(ǩ��ʱ��,��ʾʱ��)
                    End Select
                    F1Main.SetCellFormat vCell
                    Call F1Main.SetBorder(-1, .CellLineLeft, .CellLineRight, .CellLineTop, .CellLineBottom, 0, -1, .CellLineLeftColor, .CellLineRightColor, .CellLineTopColor, .CellLineBottomColor)
                End If
            End With
        Next
    End With
    mblnInit = False
End Sub
Private Sub GetTextELement(ByVal strCellKey As String)
'���ܣ�����Text Element��дF1Main�еĵ�Ԫ����������ı�
Dim i As Long, lCount As Long, strTmp As String, ltCount As Long, leCount As Long, cleTmp As cTabElement
    With Doc.Cells(strCellKey)
        ltCount = UBound(Split(.TextKey, "|")): If ltCount < 0 Then ltCount = 0
        leCount = UBound(Split(.ElementKey, "|")): If leCount < 0 Then leCount = 0
        lCount = ltCount + leCount
        For i = 1 To lCount
            Set cleTmp = .clElement(Doc.Elements, i)
            If cleTmp Is Nothing Then '�ô���Ϊ�ı�
                strTmp = strTmp & ToVarchar(.clText(Doc.Texts, i).�����ı�, 4000)
            Else
                With Doc.Elements("K" & cleTmp.Key)
                    If .�滻�� = 1 And (Doc.ET = TabET_�������༭ Or Doc.ET = TabET_���������) Then
                        If Trim(.�����ı�) = "" Then
                            If .�Զ�ת�ı� Then
                                strTmp = strTmp & " " & .Ҫ�ص�λ
                            Else
                                strTmp = strTmp & "[" & .Ҫ������ & "]" & .Ҫ�ص�λ
                            End If
                        Else
                            strTmp = strTmp & .�����ı� & .Ҫ�ص�λ
                        End If
                    Else
                        If .������̬ = 0 Then
                            strTmp = strTmp & IIf(Trim(.�����ı�) = "", "[" & .Ҫ������ & "]", .�����ı�) & .Ҫ�ص�λ
                        Else
                            strTmp = strTmp & .�����ı� & .Ҫ�ص�λ
                        End If
                    End If
                End With
            End If
        Next
        .�����ı� = strTmp
        F1Main.TextRC(.Row, .Col) = strTmp
    End With
End Sub
Private Sub PaintPictureOnTable(ByVal strCellKey As String)
'����:��ָ����Ԫ���ͼ
Dim objTmp As Object, vR As F1Rect, i As Integer, lHheight As Long, lHwidth As Long, lpLeft As Long, lpTop As Long 'ͼƬ��,����,�̶��и߶�,�̶��п��,ͼƬ��XY����
Dim lsRow As Long, leRow As Long, lsCol As Long, leCol As Long '������ֹ����
Dim lsPosX As Long, lsPosY As Long, lpHeight As Long, lpWidth As Long 'ͼƬԴ����XY����,ͼƬ�߿�

    If F1Main.ShowColHeading Then lHheight = F1Main.HdrHeight Else lHheight = 0 '�̶��и߶�
    If F1Main.ShowRowHeading Then lHwidth = F1Main.HdrWidth Else lHwidth = 0    '�̶��п��

    With Doc.Cells(strCellKey)
        If .PictureKey = "" Then Exit Sub
        If Doc.Pictures("K" & .PictureKey).OrigPic.Handle = 0 Then Exit Sub
        
        'ȷ��ͼƬ����������
        If .Merge Then  'MergeRange���ݸ�ʽ (���Ϸ�)��,��;(���·�)��,��
            lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): leRow = Split(Split(.MergeRange, ";")(1), ",")(0)
            lsCol = Split(Split(.MergeRange, ";")(0), ",")(1): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
        Else
            lsRow = .Row: leRow = .Row: lsCol = .Col: leCol = .Col
        End If
        Set vR = F1Main.RangeToTwipsEx(lsRow, lsCol, leRow, leCol)
        'ȷ��ͼƬ���С��λ�ü��ü�����
        If vR.Right - lHwidth <= 0 Or vR.Bottom - lHheight <= 0 Then '���ڿ���ʾ����
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        ElseIf vR.Left >= 0 And vR.Top >= 0 Then '�����ڱ���м�
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width: lpHeight = vR.Height: lsPosX = 0: lsPosY = 0
        ElseIf vR.Left >= 0 And vR.Top < 0 Then '�����Ϸ���������(��������)
            lpLeft = F1Main.Left + vR.Left: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = 0: lsPosY = vR.Height - lpHeight
        ElseIf vR.Left < 0 And vR.Top >= 0 Then '�����󷽲�������(��������)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + vR.Top: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height: lsPosX = vR.Width - lpWidth: lsPosY = 0
        ElseIf vR.Left < 0 And vR.Top < 0 Then '�����Ϸ��󷽶�����(��������)
            lpLeft = F1Main.Left + lHwidth: lpTop = F1Main.Top + lHheight: lpWidth = vR.Width + vR.Left - lHwidth: lpHeight = vR.Height + vR.Top - lHheight: lsPosX = vR.Width - lpWidth: lsPosY = vR.Height - lpHeight
        Else                                    '���ڿ���ʾ����
            If ChkControl(PicDy(.Index)) Then
                PicDy(.Index).Visible = False
            End If
            Exit Sub
        End If
        
        '��̬����ͼƬ������
        If Not ChkControl(PicDy(.Index)) Then
            Load PicDy(.Index)
        End If
        Set objTmp = PicDy(.Index): objTmp.Cls
        objTmp.Tag = .MergeRange & "|" & strCellKey: objTmp.ToolTipText = IIf(.�������� = cprCTReportPic, "����ͼ", "�ο�ͼ")
        objTmp.AutoRedraw = True: objTmp.BorderStyle = 0
        
        '�ȕ���ͼƬ��С��������
        LockWindowUpdate Me.hWnd
        objTmp.Width = vR.Width - Screen.TwipsPerPixelX * 2: objTmp.Height = vR.Height - Screen.TwipsPerPixelY * 2
        Set objTmp.Picture = Doc.Pictures("K" & .PictureKey).OrigPic
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height
        If .PicMarkKey <> "" Then '�б��ͼ�Ȼ���
            For i = 1 To UBound(Split(.PicMarkKey, "|"))
                ShowPicMark objTmp, Doc.PicMarks("K" & Split(.PicMarkKey, "|")(i))
            Next
        End If
        Set objTmp.Picture = objTmp.Image
        '������ʵ����ʾ��С�������ػ�
        objTmp.Move lpLeft + Screen.TwipsPerPixelX * 2, lpTop + Screen.TwipsPerPixelY * 2, lpWidth - Screen.TwipsPerPixelX * 2, lpHeight - Screen.TwipsPerPixelY * 2
        objTmp.PaintPicture objTmp.Picture, 0, 0, objTmp.Width, objTmp.Height, lsPosX, lsPosY
        objTmp.Visible = True: objTmp.ZOrder
        LockWindowUpdate 0
    End With
End Sub
