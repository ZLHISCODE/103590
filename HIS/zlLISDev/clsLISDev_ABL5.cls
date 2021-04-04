VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABL5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'ÒÇÆ÷:ÑªÆø·ÖÎöÒÇ ABL5
'³§ÉÌ:µ¤ÂóÀ×¶È
'×÷Õß:³Â¶«
'ÓÃ»§:ÖØÇì½­ÁêÒ½Ôº,ÖØÇìÖĞÉ½Ò½Ôº,ÖØÇìÊĞ¶şÔº
Private mblnStart As Boolean '¿ªÊ¼
Private mblnEnd As Boolean   '½áÊø
Private mstrData As String 'Ëû´ı½âÎö´®

Public Function GetStartCmd() As String
    'Éè±¸±ê×¼½Ó¿Ú
    '¹¦ÄÜ£º»ñÈ¡Éè±¸¿ªÊ¼·¢ËÍÃüÁî
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    'Éè±¸±ê×¼½Ó¿Ú
    '¹¦ÄÜ£º»ñÈ¡Éè±¸Í£Ö¹·¢ËÍÃüÁî
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    'Éè±¸±ê×¼½Ó¿Ú
    '¹¦ÄÜ£º½âÎöÊı¾İ
    '²ÎÊı£º
    'strSource£ºÒª½âÎöµÄÔ­Ê¼´®
    'strResult£º·µ»ØµÄ¼ìÑé½á¹û(¸÷ÒÇÆ÷½âÎö³ÌĞò±ØĞë°´ÒÔÏÂ±ê×¼×éÖ¯½á¹û)
    '   Ã¿×é¼ìÑé½á¹ûÒÔ||·Ö¸ô,ÔªËØÖ®¼äÒÔ|·Ö¸ô
    '   µÚ0¸öÔªËØ£º¼ìÑéÊ±¼ä
    '   µÚ1¸öÔªËØ£ºÑù±¾ĞòºÅ
    '   µÚ2¸öÔªËØ£º¼ìÑéÈË
    '   µÚ3¸öÔªËØ£º±ê±¾
    '   µÚ4¸öÔªËØ£ºÊÇ·ñÖÊ¿ØÆ·
    '   ´ÓµÚ5¸öÔªËØ¿ªÊ¼Îª¼ìÑé½á¹û£¬Ã¿2¸öÔªËØ±íÊ¾Ò»¸ö¼ìÑéÏîÄ¿¡£
    '       Èç£ºµÚ5i¸öÔªËØÎª¼ìÑéÏîÄ¿£¬µÚ5i+1¸öÔªËØÎª¼ìÑé½á¹û
    'strReserved£º×îºó²»ÄÜÍê³É½âÎöµÄÔ­Ê¼´®£¬ĞèÒª·µ»Ø×÷ºóĞø´¦Àí
    'strCmd£ºÈç¹ûĞèÒª£¬¿É·µ»ØÏòÉè±¸·¢ËÍµÄÃüÁî
    
    Dim lng_Stx As String, lng_Etx As Long, str_Code As String
    Dim varIn As Variant '´æ·ÅÊäÈëÊı¾İ
    Dim strLine As String
    Dim str±ê±¾ºÅ As String, strÈÕÆÚ As String, str¼ìÑéÈË±ê±¾ÖÊ¿Ø As String, strÖ¸±ê As String, strÖµ As String
    Dim strÒªÌáÈ¡Ö¸±ê As String, varÒªÌáÈ¡Ö¸±ê As Variant
    Dim iStrPoint As Integer
    
    str¼ìÑéÈË±ê±¾ÖÊ¿Ø = " |ÑªÒº|0"
    '³õÊ¼»¯·µ»ØÖµ
    strReserved = "": strResult = "": strCmd = ""
    '-- ¶¨ÒåÒªÌáÈ¡µÄÖ¸±ê
    strÒªÌáÈ¡Ö¸±ê = "Temp;FIO2;tHb;pH(T);pCO2(T);pO2(T);AaDpO2,T;AaDpO2;pH;pCO2;pO2;HCO3-;sO2;tCO2;ABE;SBC;SBE"
    varÒªÌáÈ¡Ö¸±ê = Split(strÒªÌáÈ¡Ö¸±ê, ";")
    
    
    If strSource Like "*Pc40R*Pc40R*Analyzer [#]*" Then
        strReserved = strSource
        lng_Stx = InStr(strReserved, "Pc40R")
        lng_Etx = InStr(strReserved, "Analyzer #") + 10
        str_Code = Mid(strReserved, lng_Stx, lng_Etx - lng_Stx)
        strReserved = Mid(strReserved, lng_Etx + 1)
        If strReserved = Space(Len(strReserved)) Then strReserved = ""
        
        varIn = Split(str_Code, vbNewLine)
        If UBound(varIn) <= 0 Then
            varIn = Split(str_Code, Chr(10))
        End If

        strLine = ""
        For iStrPoint = LBound(varIn) To UBound(varIn)

            strLine = varIn(iStrPoint)

            '--- ±ê±¾ºÅ
            If strLine Like "*W*#*" Then
                str±ê±¾ºÅ = Trim(Mid(strLine, InStr(strLine, "#") + 1))
                If str±ê±¾ºÅ = "" And mblnEnd Then Exit Sub
            End If

            '--- ÈÕÆÚ
            If strLine Like "Date*.*.*" Then
                strÈÕÆÚ = Format(CDate(Format(Replace(Replace(strLine, "Date                   ", ""), ".", "-"), "dd-MM-yy")), "yyyy-MM-dd")
            End If


            '-- ÍË³ö
            If InStr(strLine, "Analyzer #") > 0 And strResult = "" Then
                If strResult = "" And str±ê±¾ºÅ <> "" And strÖ¸±ê <> "" Then
                    If strÈÕÆÚ = "" Then strÈÕÆÚ = Format(Now(), "yyyy-MM-dd")
                    strResult = strÈÕÆÚ & "|" & str±ê±¾ºÅ & "|" & str¼ìÑéÈË±ê±¾ÖÊ¿Ø & strÖ¸±ê
                End If
                Call WriteLog("Abl5.Analyse ½âÂëÍê ", mstrData, strResult)
                mblnEnd = False: mblnStart = False: mstrData = ""
                Exit Sub
            End If

            '-- ÒÔÏÂÊÇÖ¸±ê "
            Dim i As Integer
            For i = LBound(varÒªÌáÈ¡Ö¸±ê) To UBound(varÒªÌáÈ¡Ö¸±ê)
                If InStr(strLine, varÒªÌáÈ¡Ö¸±ê(i)) > 0 Then
                    If varÒªÌáÈ¡Ö¸±ê(i) = "[CPc40R" Then
                        strÖ¸±ê = strÖ¸±ê & "|" & "tHb" & "|" & Trim(Replace(Replace(Mid(strLine, InStr(strLine, varÒªÌáÈ¡Ö¸±ê(i)) + 20), varÒªÌáÈ¡Ö¸±ê(i), ""), "g/L", ""))
                    Else
                        strÖµ = Trim(Replace(Mid(strLine, 1, 28), varÒªÌáÈ¡Ö¸±ê(i), ""))
                        strÖµ = Trim(Replace(strÖµ, "(T)", ""))
                        strÖ¸±ê = strÖ¸±ê & "|" & varÒªÌáÈ¡Ö¸±ê(i) & "|" & IIf(strÖµ = "", "£­£­", strÖµ)
                    End If
                    Exit For
                End If
            Next
        Next
        
    Else
        strReserved = strSource
    End If
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '½âÎöĞèÒªÏòÒÇÆ÷·¢ËÍµÄ±ê±¾ĞÅÏ¢
    '   ÔªËØÖ®¼äÒÔ|·Ö¸ô
    '   µÚ0¸öÔªËØ£º¼ìÑéÊ±¼ä
    '   µÚ1¸öÔªËØ£ºÑù±¾ĞòºÅ
    '   µÚ2¸öÔªËØ£º¼ìÑéÈË
    '   µÚ3¸öÔªËØ£º±ê±¾
    '   µÚ4¸öÔªËØ£ºÊÇ·ñÖÊ¿ØÆ·
    '   µÚ5¡«9ÔªËØ£ºÏµÍ³±£Áô
    '   ´ÓµÚ10¸öÔªËØ¿ªÊ¼ÎªĞèÒªµÄ¼ìÑéÏîÄ¿¡£
    SendSample = ""
End Function


'------------------------------------------------------------------------------------------------------------
'    Dim iStrPoint As Integer '×Ö·ûÖ¸Õë
'    Dim strBuffer As String, strData As String
'    Dim iCurrGrpEntry As Integer 'µ±Ç°Õâ×é½á¹ûµÄÈë¿Ú×Ö·ûÎ»ÖÃ
'
'    Dim varIn As Variant '´æ·ÅÊäÈëÊı¾İ
'    Dim strTime As String
'
'    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
'    Dim strLine As String
'    Dim str±ê±¾ºÅ As String, strÈÕÆÚ As String, str¼ìÑéÈË±ê±¾ÖÊ¿Ø As String, strÖ¸±ê As String, strÖµ As String
'    Dim strÒªÌáÈ¡Ö¸±ê As String, varÒªÌáÈ¡Ö¸±ê As Variant
'
'    On Error GoTo errHandle
'    str¼ìÑéÈË±ê±¾ÖÊ¿Ø = " |ÑªÒº|0"
'    '³õÊ¼»¯·µ»ØÖµ
'    strReserved = "": strResult = "": strCmd = ""
'    '-- ¶¨ÒåÒªÌáÈ¡µÄÖ¸±ê
'    strÒªÌáÈ¡Ö¸±ê = "Temp;FIO2;pH;pCO2;pO2;HCO3-;sO2;tCO2;ABE;SBC;SBE"
'    varÒªÌáÈ¡Ö¸±ê = Split(strÒªÌáÈ¡Ö¸±ê, ";")
'
'    If strSource = "" Then Exit Sub
'    Call WriteLog("Abl5.Analyse ½øÈë ", strSource, mstrData)
'    varIn = Split(strSource, vbNewLine)
'    If UBound(varIn) <= 0 Then
'        varIn = Split(strSource, Chr(10))
'    End If
'
'    If mblnStart = True And mblnEnd = False Then
'            'strReserved = strSource
'        mstrData = mstrData & strSource
'        Call WriteLog("Abl5.Analyse ÊÕÈ¡Êı¾İ ", strSource, mstrData)
'        If mstrData Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse ÕÒµ½½áÊø·û ", strSource, mstrData)
'        End If
'    End If
'
'
'    If strSource Like "*ABL5*" Then
'        mblnStart = True
'        'strReserved = strSource
'        Call WriteLog("Abl5.Analyse ÕÒµ½¿ªÊ¼·û ", strSource, mstrData)
'        If strSource Like "*Analyzer *" Then
'            mblnEnd = True
'            Call WriteLog("Abl5.Analyse È«²¿´«Èë,ÕÒµ½½áÊø·û ", strSource, mstrData)
'        End If
'        mstrData = strSource
'    End If
'
'    If mblnStart And mblnEnd And mstrData <> "" Then
'        Call WriteLog("Abl5.Analyse ¿ªÊ¼½âÂë ", strSource, mstrData)
'    End If
'    Exit Sub
'errHandle:
'    Call WriteLog("Abl5.Analyse ½âÂë´íÎó", strSource, strResult & vbNewLine & Err.Description)
