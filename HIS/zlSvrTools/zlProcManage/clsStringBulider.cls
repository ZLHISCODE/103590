VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringBulider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ËµÃ÷£º´ËÀàÀ´×Ô"¹Â·«"(2005-11-11 02:44:32),Õâ¸ö×Ö·û´®²Ù×÷ÀàµÄ×ÔÔö»º³åÇøµÄÌØµãÊÇ£ºµ±»º³åÇø²»¹»´óÊ±½«ÒÔ128£¨128*2£©µÄ´óĞ¡Ôö¼Ó»º³åÇø´óĞ¡

Private m_Buffer() As Byte   '»º³åÇø
Private m_lpBuffer As Long   '·ÖÅäµÄÄÚ´æ¿éÖ¸Õë

Private m_AllocSize As Long  'Êµ¼Ê·ÖÅäµÄÄÚ´æ´óĞ¡
Private m_Length As Long     'µ±Ç°ÒÑ¾­Ê¹ÓÃµÄÄÚ´æ´óĞ¡

'*************************************************************************
'   ÉèÖÃ»º³åÇøÈİÁ¿´óĞ¡
'*************************************************************************
Property Let Capacity(ByVal Size As Long)
    If Size < 1 Then Exit Property

    Size = Size + Size 'ÓÉÓÚvbµÄ×Ö·û´®ÊÇBSTR×Ö·ûÕ¼2¸ö×Ö½ÚËùÒÔ·ÖÅä2±¶µÄÄÚ´æ
    If Size <= m_AllocSize Then Exit Property 'ÄÚ´æ¿é»¹¹»ÓÃ

    '   ·ÖÅäÒ»¸ö»º³åÇø
    If m_AllocSize = 0 Or m_Length = 0 Then 'Ã»ÓĞÊı¾İ»ò»¹Ã»ÓĞ·ÖÅä¹ıÄÚ´æ
        ReDim m_Buffer(0 To Size)
    Else
        ReDim Preserve m_Buffer(0 To Size)
    End If
    m_lpBuffer = VarPtr(m_Buffer(0))        '»º³åÇøÖ¸Õë

    m_AllocSize = Size
End Property

'*************************************************************************
'   »ñÈ¡µ±Ç°×Ö·û´®³¤¶È
'*************************************************************************
Property Get Length() As Long
    Length = m_Length / 2
End Property

'*************************************************************************
'   Çå¿ÕÊı¾İ
'*************************************************************************
Sub Clear()
    m_Length = 0
End Sub

'*************************************************************************
'ÿ ÔÚ×Ö·û´®»º³åÇøµÄºóÃæÔö¼Ó×Ö·û´®
'   ²Î  Êı£ºÒªÔö¼ÓµÄ×Ö·û´®
'   ·µ»ØÖµ£ºÎŞ
'*************************************************************************
Sub Append(ByVal str As String)
    Const IncreaseStep = 128
    Dim nLen As Long 'ĞÂ¼ÓÈëµÄ×Ö·û´®³¤¶È

    nLen = Len(str)
    nLen = nLen + nLen

    '   Èç¹ûm_AllocSize <= m_Length + nLen£¬ÔòÒÔ 128+ nLen ¸ö×Ö·ûµÄÔöÁ¿Ôö¼ÓÄÚ´æ
    If m_AllocSize <= m_Length + nLen Then
        Me.Capacity = (m_Length + nLen) / 2 + IncreaseStep
    End If

    '   ¿½±´Êı¾İµ½»º³åÇø
    Call CopyMemory(ByVal (m_lpBuffer + m_Length), ByVal StrPtr(str), nLen)
    m_Length = m_Length + nLen
End Sub

'*************************************************************************
'   ×ª»»ÎªvbµÄ×Ö·û´®
'*************************************************************************
Property Get ToString() As String
    If m_Length > 1 Then
        ToString = Left(m_Buffer, m_Length / 2) '×¢Òâ²»ÒªÓÃLeft$
    End If
End Property

'*************************************************************************
'   Ñ°ÕÒ str ÔÚ»º³åÇøÖĞµÚÒ»´Î³öÏÖµÄÎ»ÖÃ(×Ö·ûÎ»ÖÃ´Ó 1 ÊıÆğ)
'   ²Î  Êı£ºÒª²éÕÒµÄ×Ö·û´®,[ÊÇ·ñºöÊÓ´óĞ¡Ğ´(Ä¬ÈÏºöÊÓ´óĞ¡Ğ´)]
'   ·µ»ØÖµ£º³É¹¦Ôò²»Îª-1
'*************************************************************************
Function IndexOf(ByVal str As String, Optional ByVal IsIgnoreCase As Boolean = True) As Long
    If m_Length < 1 Then
        IndexOf = -1
        Exit Function
    End If

    If IsIgnoreCase Then
        IndexOf = StrCSpnI(m_lpBuffer, StrPtr(str))
    Else
        IndexOf = StrCSpn(m_lpBuffer, StrPtr(str))
    End If

    If IndexOf >= 0 Then IndexOf = IndexOf + 1
End Function

'*************************************************************************
'   Ñ°ÕÒ str ÔÚ»º³åÇøÖĞ×îºóÒ»´Î³öÏÖµÄÎ»ÖÃ(×Ö·ûÎ»ÖÃ´Ó 1 ÊıÆğ)
'   ²Î  Êı£ºÒª²éÕÒµÄ×Ö·û´®,[ÊÇ·ñºöÊÓ´óĞ¡Ğ´(Ä¬ÈÏºöÊÓ´óĞ¡Ğ´)]
'   ·µ»ØÖµ£º³É¹¦Ôò²»Îª-1
'*************************************************************************
Function LastIndexOf(ByVal str As String, Optional ByVal IsIgnoreCase As Boolean = True) As Long
    If m_Length < 1 Then
        LastIndexOf = -1
        Exit Function
    End If

    If IsIgnoreCase Then
        LastIndexOf = StrRStrI(m_lpBuffer, m_lpBuffer + m_Length, StrPtr(str))
    Else
        LastIndexOf = StrRStr(m_lpBuffer, m_lpBuffer + m_Length, StrPtr(str))
    End If

    If LastIndexOf >= m_lpBuffer Then
        LastIndexOf = (LastIndexOf - m_lpBuffer) / 2 + 1
    Else
        LastIndexOf = -1
    End If
End Function

Private Sub Class_Initialize()
    m_AllocSize = 0
    m_Length = 0
End Sub

Private Sub Class_Terminate()
    Erase m_Buffer
End Sub
