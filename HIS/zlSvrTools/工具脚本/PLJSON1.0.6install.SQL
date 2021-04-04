Create Or Replace Type Json_Value As Object
(
  Typeval         Number(1), /* 1 = object, 2 = array, 3 = string, 4 = number, 5 = bool, 6 = null */
  Str             Varchar2(32767),
  Num             Number, /* store 1 as true, 0 as false */
  Object_Or_Array Sys.Anydata, /* object or array in here */
  Extended_Str    Clob,

 /* mapping */
  Mapname Varchar2(4000),
  Mapindx Number(32),

  Constructor Function Json_Value(Object_Or_Array Sys.Anydata) Return Self As Result,
  Constructor Function Json_Value(Str Varchar2, Esc Boolean Default True) Return Self As Result,
  Constructor Function Json_Value(Str Clob, Esc Boolean Default True) Return Self As Result,
  Constructor Function Json_Value(Num Number) Return Self As Result,
  Constructor Function Json_Value(b Boolean) Return Self As Result,
  Constructor Function Json_Value Return Self As Result,
  Static Function Makenull Return Json_Value,

  Member Function Get_Type Return Varchar2,
  Member Function Get_String(Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null) Return Varchar2,
  Member Procedure Get_String(Self In Json_Value, Buf In Out Nocopy Clob),
  Member Function Get_Number Return Number,
  Member Function Get_Bool Return Boolean,
  Member Function Get_Null Return Varchar2,

  Member Function Is_Object Return Boolean,
  Member Function Is_Array Return Boolean,
  Member Function Is_String Return Boolean,
  Member Function Is_Number Return Boolean,
  Member Function Is_Bool Return Boolean,
  Member Function Is_Null Return Boolean,

 /* Output methods */
  Member Function To_Char(Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2,
  Member Procedure To_Clob(Self In Json_Value, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True),
  Member Procedure Print(Self In Json_Value, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null), --32512 is maximum
  Member Procedure Htp(Self In Json_Value, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null),

  Member Function Value_Of(Self In Json_Value, Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null) Return Varchar2
)
Not Final;
/

Create Or Replace Type Json_Value_Array As Table Of Json_Value;
/

Create Or Replace Type Json_List As Object
(
  List_Data Json_Value_Array,
  Constructor Function Json_List Return Self As Result,
  Constructor Function Json_List(Str Varchar2) Return Self As Result,
  Constructor Function Json_List(Str Clob) Return Self As Result,
  Constructor Function Json_List(Cast Json_Value) Return Self As Result,

  Member Procedure Append(Self In Out Nocopy Json_List, Elem Json_Value, Position Pls_Integer Default Null),
  Member Procedure Append(Self In Out Nocopy Json_List, Elem Varchar2, Position Pls_Integer Default Null),
  Member Procedure Append(Self In Out Nocopy Json_List, Elem Number, Position Pls_Integer Default Null),
  Member Procedure Append(Self In Out Nocopy Json_List, Elem Boolean, Position Pls_Integer Default Null),
  Member Procedure Append(Self In Out Nocopy Json_List, Elem Json_List, Position Pls_Integer Default Null),

  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_Value),
  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Varchar2),
  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Number),
  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Boolean),
  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_List),

  Member Function Count Return Number,
  Member Procedure Remove(Self In Out Nocopy Json_List, Position Pls_Integer),
  Member Procedure Remove_First(Self In Out Nocopy Json_List),
  Member Procedure Remove_Last(Self In Out Nocopy Json_List),
  Member Function Get(Position Pls_Integer) Return Json_Value,
  Member Function Head Return Json_Value,
  Member Function Last Return Json_Value,
  Member Function Tail Return Json_List,

  /* Output methods */
  Member Function To_Char(Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2,
  Member Procedure To_Clob(Self In Json_List, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True),
  Member Procedure Print(Self In Json_List, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null), --32512 is maximum
  Member Procedure Htp(Self In Json_List, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null),

  /* json path */
  Member Function Path(Json_Path Varchar2, Base Number Default 1) Return Json_Value,
  /* json path_put */
  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_Value, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Varchar2, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Number, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Boolean, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_List, Base Number Default 1),

  /* json path_remove */
  Member Procedure Path_Remove(Self In Out Nocopy Json_List, Json_Path Varchar2, Base Number Default 1),
  Member Function To_Json_Value Return Json_Value
)
Not Final;
/

Create Or Replace Type Json As Object
(
  /* Variables */
  Json_Data           Json_Value_Array,
  Check_For_Duplicate Number,

  /* Constructors */
  Constructor Function Json Return Self As Result,
  Constructor Function Json(Str Varchar2) Return Self As Result,
  Constructor Function Json(Str In Clob) Return Self As Result,
  Constructor Function Json(Cast Json_Value) Return Self As Result,
  Constructor Function Json(l In Out Nocopy Json_List) Return Self As Result,

  /* Member setter methods */
  Member Procedure Remove(Pair_Name Varchar2),
  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_Value, Position Pls_Integer Default Null),
  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Varchar2, Position Pls_Integer Default Null),
  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Number, Position Pls_Integer Default Null),
  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Boolean, Position Pls_Integer Default Null),
  Member Procedure Check_Duplicate(Self In Out Nocopy Json, v_Set Boolean),
  Member Procedure Remove_Duplicates(Self In Out Nocopy Json),

  /* deprecated putter use json_value */
  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json, Position Pls_Integer Default Null),
  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_List, Position Pls_Integer Default Null),

  /* Member getter methods */
  Member Function Count Return Number,
  Member Function Get(Pair_Name Varchar2) Return Json_Value,
  Member Function Get(Position Pls_Integer) Return Json_Value,
  Member Function Index_Of(Pair_Name Varchar2) Return Number,
  Member Function Exist(Pair_Name Varchar2) Return Boolean,

  /* Output methods */
  Member Function To_Char(Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2,
  Member Procedure To_Clob(Self In Json, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True),
  Member Procedure Print(Self In Json, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null), --32512 is maximum
  Member Procedure Htp(Self In Json, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null),

  Member Function To_Json_Value Return Json_Value,
  /* json path */
  Member Function Path(Json_Path Varchar2, Base Number Default 1) Return Json_Value,

  /* json path_put */
  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_Value, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Varchar2, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Number, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Boolean, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_List, Base Number Default 1),
  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Json, Base Number Default 1),

  /* json path_remove */
  Member Procedure Path_Remove(Self In Out Nocopy Json, Json_Path Varchar2, Base Number Default 1),

  /* map functions */
  Member Function Get_Values Return Json_List,
  Member Function Get_Keys Return Json_List
)
Not Final;
/

Create Or Replace Package Json_Parser As
  Type Rtoken Is Record(
    Type_Name     Varchar2(7),
    Line          Pls_Integer,
    Col           Pls_Integer,
    Data          Varchar2(32767),
    Data_Overflow Clob); -- max_string_size

  Type Ltokens Is Table Of Rtoken Index By Pls_Integer;
  Type Json_Src Is Record(
    Len    Number,
    Offset Number,
    Src    Varchar2(32767),
    s_Clob Clob);

  Json_Strict Boolean Not Null := False;

  Function Next_Char(Indx Number, s In Out Nocopy Json_Src) Return Varchar2;
  Function Next_Char2(Indx Number, s In Out Nocopy Json_Src, Amount Number Default 1) Return Varchar2;

  Function Prepareclob(Buf In Clob) Return Json_Parser.Json_Src;
  Function Preparevarchar2(Buf In Varchar2) Return Json_Parser.Json_Src;
  Function Lexer(Jsrc In Out Nocopy Json_Src) Return Ltokens;
  Procedure Print_Token(t Rtoken);

  Function Parser(Str Varchar2) Return Json;
  Function Parse_List(Str Varchar2) Return Json_List;
  Function Parse_Any(Str Varchar2) Return Json_Value;
  Function Parser(Str Clob) Return Json;
  Function Parse_List(Str Clob) Return Json_List;
  Function Parse_Any(Str Clob) Return Json_Value;
  Procedure Remove_Duplicates(Obj In Out Nocopy Json);
  Function Get_Version Return Varchar2;

End Json_Parser;
/



Create Or Replace Package Body Json_Parser As

  Decimalpoint Varchar2(1 Char) := '.';

  Procedure Updatedecimalpoint As
  Begin
    Select Substr(Value, 1, 1)
    Into Decimalpoint
    From Nls_Session_Parameters
    Where Parameter = 'NLS_NUMERIC_CHARACTERS';
  End Updatedecimalpoint;
  Function Next_Char(Indx Number, s In Out Nocopy Json_Src) Return Varchar2 As
  Begin
    If (Indx > s.Len) Then
      Return Null;
    End If;
    If (Indx > 4000 + s.Offset Or Indx <= s.Offset) Then
      s.Offset := Indx - (Indx Mod 4000);
      If s.Offset = Indx Then
        s.Offset := s.Offset - 4000;
      End If;
      s.Src := Dbms_Lob.Substr(s.s_Clob, 4000, s.Offset + 1);
    End If;
    Return Substr(s.Src, Indx - s.Offset, 1);
  End;

  Function Next_Char2(Indx Number, s In Out Nocopy Json_Src, Amount Number Default 1) Return Varchar2 As
    Buf Varchar2(32767) := '';
  Begin
    For i In 1 .. Amount Loop
      Buf := Buf || Next_Char(Indx - 1 + i, s);
    End Loop;
    Return Buf;
  End;

  Function Prepareclob(Buf Clob) Return Json_Parser.Json_Src As
    Temp Json_Parser.Json_Src;
  Begin
    Temp.s_Clob := Buf;
    Temp.Offset := 0;
    Temp.Src    := Dbms_Lob.Substr(Buf, 4000, Temp.Offset + 1);
    Temp.Len    := Dbms_Lob.Getlength(Buf);
    Return Temp;
  End;

  Function Preparevarchar2(Buf Varchar2) Return Json_Parser.Json_Src As
    Temp Json_Parser.Json_Src;
  Begin
    Temp.s_Clob := Buf;
    Temp.Offset := 0;
    Temp.Src    := Substr(Buf, 1, 4000);
    Temp.Len    := Length(Buf);
    Return Temp;
  End;

  Procedure Debug(Text Varchar2) As
  Begin
    Dbms_Output.Put_Line(Text);
  End;

  Procedure Print_Token(t Rtoken) As
  Begin
    Dbms_Output.Put_Line('Line: ' || t.Line || ' - Column: ' || t.Col || ' - Type: ' || t.Type_Name || ' - Content: ' ||
                         t.Data);
  End Print_Token;

  /* SCANNER FUNCTIONS START */
  Procedure s_Error(Text Varchar2, Line Number, Col Number) As
  Begin
    Raise_Application_Error(-20100, 'JSON Scanner exception @ line: ' || Line || ' column: ' || Col || ' - ' || Text);
  End;

  Procedure s_Error(Text Varchar2, Tok Rtoken) As
  Begin
    Raise_Application_Error(-20100, 'JSON Scanner exception @ line: ' || Tok.Line || ' column: ' || Tok.Col || ' - ' || Text);
  End;

  Function Mt(t Varchar2, l Pls_Integer, c Pls_Integer, d Varchar2) Return Rtoken As
    Token Rtoken;
  Begin
    Token.Type_Name := t;
    Token.Line      := l;
    Token.Col       := c;
    Token.Data      := d;
    Return Token;
  End;

  Function Lexnumber(Jsrc In Out Nocopy Json_Src, Tok In Out Nocopy Rtoken, Indx In Out Nocopy Pls_Integer)
    Return Pls_Integer As
    Numbuf    Varchar2(4000) := '';
    Buf       Varchar2(4);
    Checkloop Boolean;
  Begin
    Buf := Next_Char(Indx, Jsrc);
    If (Buf = '-') Then
      Numbuf := '-';
      Indx   := Indx + 1;
    End If;
    Buf := Next_Char(Indx, Jsrc);
    --0 or [1-9]([0-9])
    If (Buf = '0') Then
      Numbuf := Numbuf || '0';
      Indx   := Indx + 1;
      Buf    := Next_Char(Indx, Jsrc);
    Elsif (Buf >= '1' And Buf <= '9') Then
      Numbuf := Numbuf || Buf;
      Indx   := Indx + 1;
      --read digits
      Buf := Next_Char(Indx, Jsrc);
      While (Buf >= '0' And Buf <= '9') Loop
        Numbuf := Numbuf || Buf;
        Indx   := Indx + 1;
        Buf    := Next_Char(Indx, Jsrc);
      End Loop;
    End If;
    --fraction
    If (Buf = '.') Then
      Numbuf    := Numbuf || Buf;
      Indx      := Indx + 1;
      Buf       := Next_Char(Indx, Jsrc);
      Checkloop := False;
      While (Buf >= '0' And Buf <= '9') Loop
        Checkloop := True;
        Numbuf    := Numbuf || Buf;
        Indx      := Indx + 1;
        Buf       := Next_Char(Indx, Jsrc);
      End Loop;
      If (Not Checkloop) Then
        s_Error('Expected: digits in fraction', Tok);
      End If;
    End If;
    --exp part
    If (Buf In ('e', 'E')) Then
      Numbuf := Numbuf || Buf;
      Indx   := Indx + 1;
      Buf    := Next_Char(Indx, Jsrc);
      If (Buf = '+' Or Buf = '-') Then
        Numbuf := Numbuf || Buf;
        Indx   := Indx + 1;
        Buf    := Next_Char(Indx, Jsrc);
      End If;
      Checkloop := False;
      While (Buf >= '0' And Buf <= '9') Loop
        Checkloop := True;
        Numbuf    := Numbuf || Buf;
        Indx      := Indx + 1;
        Buf       := Next_Char(Indx, Jsrc);
      End Loop;
      If (Not Checkloop) Then
        s_Error('Expected: digits in exp', Tok);
      End If;
    End If;
  
    Tok.Data := Numbuf;
    Return Indx;
  End Lexnumber;

  -- [a-zA-Z]([a-zA-Z0-9])*
  Function Lexname(Jsrc In Out Nocopy Json_Src, Tok In Out Nocopy Rtoken, Indx In Out Nocopy Pls_Integer)
    Return Pls_Integer As
    Varbuf Varchar2(32767) := '';
    Buf    Varchar(4);
    Num    Number;
  Begin
    Buf := Next_Char(Indx, Jsrc);
    While (Regexp_Like(Buf, '^[[:alnum:]\_]$', 'i')) Loop
      Varbuf := Varbuf || Buf;
      Indx   := Indx + 1;
      Buf    := Next_Char(Indx, Jsrc);
      If (Buf Is Null) Then
        Goto Retname;
        --debug('Premature string ending');
      End If;
    End Loop;
    <<retname>>
    Tok.Data := Varbuf;
    Return Indx - 1;
  End Lexname;

  Procedure Updateclob(v_Extended In Out Nocopy Clob, v_Str Varchar2) As
  Begin
    Dbms_Lob.Writeappend(v_Extended, Length(v_Str), v_Str);
  End Updateclob;

  Function Lexstring(Jsrc In Out Nocopy Json_Src, Tok In Out Nocopy Rtoken, Indx In Out Nocopy Pls_Integer, Endchar Char)
    Return Pls_Integer As
    v_Extended Clob := Null;
    v_Count    Number := 0;
    Varbuf     Varchar2(32767) := '';
    Buf        Varchar(4);
    Wrong      Boolean;
  Begin
    Indx := Indx + 1;
    Buf  := Next_Char(Indx, Jsrc);
    While (Buf != Endchar) Loop
      --clob control
      If (v_Count > 8191) Then
        --crazy oracle error (16383 is the highest working length with unistr - 8192 choosen to be safe)
        If (v_Extended Is Null) Then
          v_Extended := Empty_Clob();
          Dbms_Lob.Createtemporary(v_Extended, True);
        End If;
        Updateclob(v_Extended, Unistr(Varbuf));
        Varbuf  := '';
        v_Count := 0;
      End If;
      If (Buf = Chr(13) Or Buf = Chr(9) Or Buf = Chr(10)) Then
        s_Error('Control characters not allowed (CHR(9),CHR(10)CHR(13))', Tok);
      End If;
      If (Buf = '\') Then
        --varbuf := varbuf || buf;
        Indx := Indx + 1;
        Buf  := Next_Char(Indx, Jsrc);
        Case
          When Buf In ('\') Then
            Varbuf  := Varbuf || Buf || Buf;
            v_Count := v_Count + 2;
            Indx    := Indx + 1;
            Buf     := Next_Char(Indx, Jsrc);
          When Buf In ('"', '/') Then
            Varbuf  := Varbuf || Buf;
            v_Count := v_Count + 1;
            Indx    := Indx + 1;
            Buf     := Next_Char(Indx, Jsrc);
          When Buf = '''' Then
            If (Json_Strict = False) Then
              Varbuf  := Varbuf || Buf;
              v_Count := v_Count + 1;
              Indx    := Indx + 1;
              Buf     := Next_Char(Indx, Jsrc);
            Else
              s_Error('strictmode - expected: " \ / b f n r t u ', Tok);
            End If;
          When Buf In ('b', 'f', 'n', 'r', 't') Then
            --backspace b = U+0008
            --formfeed  f = U+000C
            --newline   n = U+000A
            --carret    r = U+000D
            --tabulator t = U+0009
            Case Buf
              When 'b' Then
                Varbuf := Varbuf || Chr(8);
              When 'f' Then
                Varbuf := Varbuf || Chr(12);
              When 'n' Then
                Varbuf := Varbuf || Chr(10);
              When 'r' Then
                Varbuf := Varbuf || Chr(13);
              When 't' Then
                Varbuf := Varbuf || Chr(9);
            End Case;
            --varbuf := varbuf || buf;
            v_Count := v_Count + 1;
            Indx    := Indx + 1;
            Buf     := Next_Char(Indx, Jsrc);
          When Buf = 'u' Then
            --four hexidecimal chars
            Declare
              Four Varchar2(4);
            Begin
              Four  := Next_Char2(Indx + 1, Jsrc, 4);
              Wrong := False;
              If (Upper(Substr(Four, 1, 1)) Not In
                 ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'a', 'b', 'c', 'd', 'e', 'f')) Then
                Wrong := True;
              End If;
              If (Upper(Substr(Four, 2, 1)) Not In
                 ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'a', 'b', 'c', 'd', 'e', 'f')) Then
                Wrong := True;
              End If;
              If (Upper(Substr(Four, 3, 1)) Not In
                 ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'a', 'b', 'c', 'd', 'e', 'f')) Then
                Wrong := True;
              End If;
              If (Upper(Substr(Four, 4, 1)) Not In
                 ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'a', 'b', 'c', 'd', 'e', 'f')) Then
                Wrong := True;
              End If;
              If (Wrong) Then
                s_Error('expected: " \u([0-9][A-F]){4}', Tok);
              End If;
              --              varbuf := varbuf || buf || four;
              Varbuf  := Varbuf || '\' || Four; --chr(to_number(four,'XXXX'));
              v_Count := v_Count + 5;
              Indx    := Indx + 5;
              Buf     := Next_Char(Indx, Jsrc);
            End;
          Else
            s_Error('expected: " \ / b f n r t u ', Tok);
        End Case;
      Else
        Varbuf  := Varbuf || Buf;
        v_Count := v_Count + 1;
        Indx    := Indx + 1;
        Buf     := Next_Char(Indx, Jsrc);
      End If;
    End Loop;
  
    If (Buf Is Null) Then
      s_Error('string ending not found', Tok);
    End If;
  
    If (v_Extended Is Not Null) Then
      Updateclob(v_Extended, Unistr(Varbuf));
      Tok.Data_Overflow := v_Extended;
      Tok.Data          := Dbms_Lob.Substr(v_Extended, 1, 32767);
    Else
      Tok.Data := Unistr(Varbuf);
    End If;
    Return Indx;
  End Lexstring;

  Function Lexer(Jsrc In Out Nocopy Json_Src) Return Ltokens As
    Tokens   Ltokens;
    Indx     Pls_Integer := 1;
    Tok_Indx Pls_Integer := 1;
    Buf      Varchar2(4);
    Lin_No   Number := 1;
    Col_No   Number := 0;
  Begin
    While (Indx <= Jsrc.Len) Loop
      --read into buf
      Buf    := Next_Char(Indx, Jsrc);
      Col_No := Col_No + 1;
      --convert to switch case
      Case
        When Buf = '{' Then
          Tokens(Tok_Indx) := Mt('{', Lin_No, Col_No, Null);
          Tok_Indx := Tok_Indx + 1;
        When Buf = '}' Then
          Tokens(Tok_Indx) := Mt('}', Lin_No, Col_No, Null);
          Tok_Indx := Tok_Indx + 1;
        When Buf = ',' Then
          Tokens(Tok_Indx) := Mt(',', Lin_No, Col_No, Null);
          Tok_Indx := Tok_Indx + 1;
        When Buf = ':' Then
          Tokens(Tok_Indx) := Mt(':', Lin_No, Col_No, Null);
          Tok_Indx := Tok_Indx + 1;
        When Buf = '[' Then
          Tokens(Tok_Indx) := Mt('[', Lin_No, Col_No, Null);
          Tok_Indx := Tok_Indx + 1;
        When Buf = ']' Then
          Tokens(Tok_Indx) := Mt(']', Lin_No, Col_No, Null);
          Tok_Indx := Tok_Indx + 1;
        When Buf = 't' Then
          If (Next_Char2(Indx, Jsrc, 4) != 'true') Then
            If (Json_Strict = False And Regexp_Like(Buf, '^[[:alpha:]]$', 'i')) Then
              Tokens(Tok_Indx) := Mt('STRING', Lin_No, Col_No, Null);
              Indx := Lexname(Jsrc, Tokens(Tok_Indx), Indx);
              Col_No := Col_No + Length(Tokens(Tok_Indx).Data) + 1;
              Tok_Indx := Tok_Indx + 1;
            Else
              s_Error('Expected: ''true''', Lin_No, Col_No);
            End If;
          Else
            Tokens(Tok_Indx) := Mt('TRUE', Lin_No, Col_No, Null);
            Tok_Indx := Tok_Indx + 1;
            Indx := Indx + 3;
            Col_No := Col_No + 3;
          End If;
        When Buf = 'n' Then
          If (Next_Char2(Indx, Jsrc, 4) != 'null') Then
            If (Json_Strict = False And Regexp_Like(Buf, '^[[:alpha:]]$', 'i')) Then
              Tokens(Tok_Indx) := Mt('STRING', Lin_No, Col_No, Null);
              Indx := Lexname(Jsrc, Tokens(Tok_Indx), Indx);
              Col_No := Col_No + Length(Tokens(Tok_Indx).Data) + 1;
              Tok_Indx := Tok_Indx + 1;
            Else
              s_Error('Expected: ''null''', Lin_No, Col_No);
            End If;
          Else
            Tokens(Tok_Indx) := Mt('NULL', Lin_No, Col_No, Null);
            Tok_Indx := Tok_Indx + 1;
            Indx := Indx + 3;
            Col_No := Col_No + 3;
          End If;
        When Buf = 'f' Then
          If (Next_Char2(Indx, Jsrc, 5) != 'false') Then
            If (Json_Strict = False And Regexp_Like(Buf, '^[[:alpha:]]$', 'i')) Then
              Tokens(Tok_Indx) := Mt('STRING', Lin_No, Col_No, Null);
              Indx := Lexname(Jsrc, Tokens(Tok_Indx), Indx);
              Col_No := Col_No + Length(Tokens(Tok_Indx).Data) + 1;
              Tok_Indx := Tok_Indx + 1;
            Else
              s_Error('Expected: ''false''', Lin_No, Col_No);
            End If;
          Else
            Tokens(Tok_Indx) := Mt('FALSE', Lin_No, Col_No, Null);
            Tok_Indx := Tok_Indx + 1;
            Indx := Indx + 4;
            Col_No := Col_No + 4;
          End If;
          -- 9 = TAB, 10 = \n, 13 = \r (Linux = \n, Windows = \r\n, Mac = \r 
        When (Buf = Chr(10)) Then
          --linux newlines
          Lin_No := Lin_No + 1;
          Col_No := 0;
        
        When (Buf = Chr(13)) Then
          --Windows or Mac way
          Lin_No := Lin_No + 1;
          Col_No := 0;
          If (Jsrc.Len >= Indx + 1) Then
            -- better safe than sorry
            Buf := Next_Char(Indx + 1, Jsrc);
            If (Buf = Chr(10)) Then
              --\r\n
              Indx := Indx + 1;
            End If;
          End If;
        
        When (Buf = Chr(9)) Then
          Null; --tabbing
        When (Buf In ('-', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')) Then
          --number
          Tokens(Tok_Indx) := Mt('NUMBER', Lin_No, Col_No, Null);
          Indx := Lexnumber(Jsrc, Tokens(Tok_Indx), Indx) - 1;
          Col_No := Col_No + Length(Tokens(Tok_Indx).Data);
          Tok_Indx := Tok_Indx + 1;
        When Buf = '"' Then
          --number
          Tokens(Tok_Indx) := Mt('STRING', Lin_No, Col_No, Null);
          Indx := Lexstring(Jsrc, Tokens(Tok_Indx), Indx, '"');
          Col_No := Col_No + Length(Tokens(Tok_Indx).Data) + 1;
          Tok_Indx := Tok_Indx + 1;
        When Buf = '''' And Json_Strict = False Then
          --number
          Tokens(Tok_Indx) := Mt('STRING', Lin_No, Col_No, Null);
          Indx := Lexstring(Jsrc, Tokens(Tok_Indx), Indx, '''');
          Col_No := Col_No + Length(Tokens(Tok_Indx).Data) + 1; --hovsa her
          Tok_Indx := Tok_Indx + 1;
        When Json_Strict = False And Regexp_Like(Buf, '^[[:alpha:]]$', 'i') Then
          Tokens(Tok_Indx) := Mt('STRING', Lin_No, Col_No, Null);
          Indx := Lexname(Jsrc, Tokens(Tok_Indx), Indx);
          If (Tokens(Tok_Indx).Data_Overflow Is Not Null) Then
            Col_No := Col_No + Dbms_Lob.Getlength(Tokens(Tok_Indx).Data_Overflow) + 1;
          Else
            Col_No := Col_No + Length(Tokens(Tok_Indx).Data) + 1;
          End If;
          Tok_Indx := Tok_Indx + 1;
        When Json_Strict = False And Buf || Next_Char(Indx + 1, Jsrc) = '/*' Then
          --strip comments
          Declare
            Saveindx Number := Indx;
            Un_Esc   Clob;
          Begin
            Indx := Indx + 1;
            Loop
              Indx := Indx + 1;
              Buf  := Next_Char(Indx, Jsrc) || Next_Char(Indx + 1, Jsrc);
              Exit When Buf = '*/';
              Exit When Buf Is Null;
            End Loop;
          
            If (Indx = Saveindx + 2) Then
              --enter unescaped mode
              --dbms_output.put_line('Entering unescaped mode');
              Un_Esc := Empty_Clob();
              Dbms_Lob.Createtemporary(Un_Esc, True);
              Indx := Indx + 1;
              Loop
                Indx := Indx + 1;
                Buf  := Next_Char(Indx, Jsrc) || Next_Char(Indx + 1, Jsrc) || Next_Char(Indx + 2, Jsrc) ||
                        Next_Char(Indx + 3, Jsrc);
                Exit When Buf = '/**/';
                If Buf Is Null Then
                  s_Error('Unexpected sequence /**/ to end unescaped data: ' || Buf, Lin_No, Col_No);
                End If;
                Buf := Next_Char(Indx, Jsrc);
                Dbms_Lob.Writeappend(Un_Esc, Length(Buf), Buf);
              End Loop;
              Tokens(Tok_Indx) := Mt('ESTRING', Lin_No, Col_No, Null);
              Tokens(Tok_Indx).Data_Overflow := Un_Esc;
              Col_No := Col_No + Dbms_Lob.Getlength(Un_Esc) + 1; --note: line count wont work properly
              Tok_Indx := Tok_Indx + 1;
              Indx := Indx + 2;
            End If;
          
            Indx := Indx + 1;
          End;
        When Buf = ' ' Then
          Null; --space
        Else
          s_Error('Unexpected char: ' || Buf, Lin_No, Col_No);
      End Case;
    
      Indx := Indx + 1;
    End Loop;
  
    Return Tokens;
  End Lexer;

  /* SCANNER END */

  /* PARSER FUNCTIONS START*/
  Procedure p_Error(Text Varchar2, Tok Rtoken) As
  Begin
    Raise_Application_Error(-20101, 'JSON Parser exception @ line: ' || Tok.Line || ' column: ' || Tok.Col || ' - ' || Text);
  End;

  Function Parseobj(Tokens Ltokens, Indx In Out Nocopy Pls_Integer) Return Json;

  Function Parsearr(Tokens Ltokens, Indx In Out Nocopy Pls_Integer) Return Json_List As
    e_Arr    Json_Value_Array := Json_Value_Array();
    Ret_List Json_List := Json_List();
    v_Count  Number := 0;
    Tok      Rtoken;
  Begin
    --value, value, value ]
    If (Indx > Tokens.Count) Then
      p_Error('more elements in array was excepted', Tok);
    End If;
    Tok := Tokens(Indx);
    While (Tok.Type_Name != ']') Loop
      e_Arr.Extend;
      v_Count := v_Count + 1;
      Case Tok.Type_Name
        When 'TRUE' Then
          e_Arr(v_Count) := Json_Value(True);
        When 'FALSE' Then
          e_Arr(v_Count) := Json_Value(False);
        When 'NULL' Then
          e_Arr(v_Count) := Json_Value;
        When 'STRING' Then
          e_Arr(v_Count) := Case
                              When Tok.Data_Overflow Is Not Null Then
                               Json_Value(Tok.Data_Overflow)
                              Else
                               Json_Value(Tok.Data)
                            End;
        When 'ESTRING' Then
          e_Arr(v_Count) := Json_Value(Tok.Data_Overflow, False);
        When 'NUMBER' Then
          e_Arr(v_Count) := Json_Value(To_Number(Replace(Tok.Data, '.', Decimalpoint)));
        When '[' Then
          Declare
            e_List Json_List;
          Begin
            Indx := Indx + 1;
            e_List := Parsearr(Tokens, Indx);
            e_Arr(v_Count) := e_List.To_Json_Value;
          End;
        When '{' Then
          Indx := Indx + 1;
          e_Arr(v_Count) := Parseobj(Tokens, Indx).To_Json_Value;
        Else
          p_Error('Expected a value', Tok);
      End Case;
      Indx := Indx + 1;
      If (Indx > Tokens.Count) Then
        p_Error('] not found', Tok);
      End If;
      Tok := Tokens(Indx);
      If (Tok.Type_Name = ',') Then
        --advance
        Indx := Indx + 1;
        If (Indx > Tokens.Count) Then
          p_Error('more elements in array was excepted', Tok);
        End If;
        Tok := Tokens(Indx);
        If (Tok.Type_Name = ']') Then
          --premature exit
          p_Error('Premature exit in array', Tok);
        End If;
      Elsif (Tok.Type_Name != ']') Then
        --error
        p_Error('Expected , or ]', Tok);
      End If;
    
    End Loop;
    Ret_List.List_Data := e_Arr;
    Return Ret_List;
  End Parsearr;

  Function Parsemem(Tokens Ltokens, Indx In Out Pls_Integer, Mem_Name Varchar2, Mem_Indx Number) Return Json_Value As
    Mem Json_Value;
    Tok Rtoken;
  Begin
    Tok := Tokens(Indx);
    Case Tok.Type_Name
      When 'TRUE' Then
        Mem := Json_Value(True);
      When 'FALSE' Then
        Mem := Json_Value(False);
      When 'NULL' Then
        Mem := Json_Value;
      When 'STRING' Then
        Mem := Case
                 When Tok.Data_Overflow Is Not Null Then
                  Json_Value(Tok.Data_Overflow)
                 Else
                  Json_Value(Tok.Data)
               End;
      When 'ESTRING' Then
        Mem := Json_Value(Tok.Data_Overflow, False);
      When 'NUMBER' Then
        Mem := Json_Value(To_Number(Replace(Tok.Data, '.', Decimalpoint)));
      When '[' Then
        Declare
          e_List Json_List;
        Begin
          Indx   := Indx + 1;
          e_List := Parsearr(Tokens, Indx);
          Mem    := e_List.To_Json_Value;
        End;
      When '{' Then
        Indx := Indx + 1;
        Mem  := Parseobj(Tokens, Indx).To_Json_Value;
      Else
        p_Error('Found ' || Tok.Type_Name, Tok);
    End Case;
    Mem.Mapname := Mem_Name;
    Mem.Mapindx := Mem_Indx;
  
    Indx := Indx + 1;
    Return Mem;
  End Parsemem;

  Function Parseobj(Tokens Ltokens, Indx In Out Nocopy Pls_Integer) Return Json As
    Type Memmap Is Table Of Number Index By Varchar2(4000); -- ive read somewhere that this is not possible - but it is!
    Mymap         Memmap;
    Nullelemfound Boolean := False;
  
    Obj      Json;
    Tok      Rtoken;
    Mem_Name Varchar(4000);
    Arr      Json_Value_Array := Json_Value_Array();
  Begin
    --what to expect
    While (Indx <= Tokens.Count) Loop
      Tok := Tokens(Indx);
      --debug('E: '||tok.type_name);
      Case Tok.Type_Name
        When 'STRING' Then
          --member
          Mem_Name := Substr(Tok.Data, 1, 4000);
          Begin
            If (Mem_Name Is Null) Then
              If (Nullelemfound) Then
                p_Error('Duplicate empty member: ', Tok);
              Else
                Nullelemfound := True;
              End If;
            Elsif (Mymap(Mem_Name) Is Not Null) Then
              p_Error('Duplicate member name: ' || Mem_Name, Tok);
            End If;
          Exception
            When No_Data_Found Then
              Mymap(Mem_Name) := 1;
          End;
        
          Indx := Indx + 1;
          If (Indx > Tokens.Count) Then
            p_Error('Unexpected end of input', Tok);
          End If;
          Tok  := Tokens(Indx);
          Indx := Indx + 1;
          If (Indx > Tokens.Count) Then
            p_Error('Unexpected end of input', Tok);
          End If;
          If (Tok.Type_Name = ':') Then
            --parse
            Declare
              Jmb Json_Value;
              x   Number;
            Begin
              x   := Arr.Count + 1;
              Jmb := Parsemem(Tokens, Indx, Mem_Name, x);
              Arr.Extend;
              Arr(x) := Jmb;
            End;
          Else
            p_Error('Expected '':''', Tok);
          End If;
          --move indx forward if ',' is found
          If (Indx > Tokens.Count) Then
            p_Error('Unexpected end of input', Tok);
          End If;
        
          Tok := Tokens(Indx);
          If (Tok.Type_Name = ',') Then
            --debug('found ,');
            Indx := Indx + 1;
            Tok  := Tokens(Indx);
            If (Tok.Type_Name = '}') Then
              --premature exit
              p_Error('Premature exit in json object', Tok);
            End If;
          Elsif (Tok.Type_Name != '}') Then
            p_Error('A comma seperator is probably missing', Tok);
          End If;
        When '}' Then
          Obj           := Json();
          Obj.Json_Data := Arr;
          Return Obj;
        Else
          p_Error('Expected string or }', Tok);
      End Case;
    End Loop;
  
    p_Error('} not found', Tokens(Indx - 1));
  
    Return Obj;
  
  End;

  Function Parser(Str Varchar2) Return Json As
    Tokens Ltokens;
    Obj    Json;
    Indx   Pls_Integer := 1;
    Jsrc   Json_Src;
  Begin
    Updatedecimalpoint();
    Jsrc   := Preparevarchar2(Str);
    Tokens := Lexer(Jsrc);
    If (Tokens(Indx).Type_Name = '{') Then
      Indx := Indx + 1;
      Obj  := Parseobj(Tokens, Indx);
    Else
      Raise_Application_Error(-20101, 'JSON Parser exception - no { start found');
    End If;
    If (Tokens.Count != Indx) Then
      p_Error('} should end the JSON object', Tokens(Indx));
    End If;
  
    Return Obj;
  End Parser;

  Function Parse_List(Str Varchar2) Return Json_List As
    Tokens Ltokens;
    Obj    Json_List;
    Indx   Pls_Integer := 1;
    Jsrc   Json_Src;
  Begin
    Updatedecimalpoint();
    Jsrc   := Preparevarchar2(Str);
    Tokens := Lexer(Jsrc);
    If (Tokens(Indx).Type_Name = '[') Then
      Indx := Indx + 1;
      Obj  := Parsearr(Tokens, Indx);
    Else
      Raise_Application_Error(-20101, 'JSON List Parser exception - no [ start found');
    End If;
    If (Tokens.Count != Indx) Then
      p_Error('] should end the JSON List object', Tokens(Indx));
    End If;
  
    Return Obj;
  End Parse_List;

  Function Parse_List(Str Clob) Return Json_List As
    Tokens Ltokens;
    Obj    Json_List;
    Indx   Pls_Integer := 1;
    Jsrc   Json_Src;
  Begin
    Updatedecimalpoint();
    Jsrc   := Prepareclob(Str);
    Tokens := Lexer(Jsrc);
    If (Tokens(Indx).Type_Name = '[') Then
      Indx := Indx + 1;
      Obj  := Parsearr(Tokens, Indx);
    Else
      Raise_Application_Error(-20101, 'JSON List Parser exception - no [ start found');
    End If;
    If (Tokens.Count != Indx) Then
      p_Error('] should end the JSON List object', Tokens(Indx));
    End If;
  
    Return Obj;
  End Parse_List;

  Function Parser(Str Clob) Return Json As
    Tokens Ltokens;
    Obj    Json;
    Indx   Pls_Integer := 1;
    Jsrc   Json_Src;
  Begin
    Updatedecimalpoint();
    --dbms_output.put_line('Using clob');
    Jsrc   := Prepareclob(Str);
    Tokens := Lexer(Jsrc);
    If (Tokens(Indx).Type_Name = '{') Then
      Indx := Indx + 1;
      Obj  := Parseobj(Tokens, Indx);
    Else
      Raise_Application_Error(-20101, 'JSON Parser exception - no { start found');
    End If;
    If (Tokens.Count != Indx) Then
      p_Error('} should end the JSON object', Tokens(Indx));
    End If;
  
    Return Obj;
  End Parser;

  Function Parse_Any(Str Varchar2) Return Json_Value As
    Tokens Ltokens;
    Obj    Json_List;
    Ret    Json_Value;
    Indx   Pls_Integer := 1;
    Jsrc   Json_Src;
  Begin
    Updatedecimalpoint();
    Jsrc := Preparevarchar2(Str);
    Tokens := Lexer(Jsrc);
    Tokens(Tokens.Count + 1).Type_Name := ']';
    Obj := Parsearr(Tokens, Indx);
    If (Tokens.Count != Indx) Then
      p_Error('] should end the JSON List object', Tokens(Indx));
    End If;
  
    Return Obj.Head();
  End Parse_Any;

  Function Parse_Any(Str Clob) Return Json_Value As
    Tokens Ltokens;
    Obj    Json_List;
    Indx   Pls_Integer := 1;
    Jsrc   Json_Src;
  Begin
    Jsrc := Prepareclob(Str);
    Tokens := Lexer(Jsrc);
    Tokens(Tokens.Count + 1).Type_Name := ']';
    Obj := Parsearr(Tokens, Indx);
    If (Tokens.Count != Indx) Then
      p_Error('] should end the JSON List object', Tokens(Indx));
    End If;
  
    Return Obj.Head();
  End Parse_Any;

  /* last entry is the one to keep */
  Procedure Remove_Duplicates(Obj In Out Nocopy Json) As
    Type Memberlist Is Table Of Json_Value Index By Varchar2(4000);
    Members       Memberlist;
    Nulljsonvalue Json_Value := Null;
    Validated     Json := Json();
    Indx          Varchar2(4000);
  Begin
    For i In 1 .. Obj.Count Loop
      If (Obj.Get(i).Mapname Is Null) Then
        Nulljsonvalue := Obj.Get(i);
      Else
        Members(Obj.Get(i).Mapname) := Obj.Get(i);
      End If;
    End Loop;
  
    Validated.Check_Duplicate(False);
    Indx := Members.First;
    Loop
      Exit When Indx Is Null;
      Validated.Put(Indx, Members(Indx));
      Indx := Members.Next(Indx);
    End Loop;
    If (Nulljsonvalue Is Not Null) Then
      Validated.Put('', Nulljsonvalue);
    End If;
  
    Validated.Check_For_Duplicate := Obj.Check_For_Duplicate;
  
    Obj := Validated;
  End;
  Function Get_Version Return Varchar2 As
  Begin
    Return 'PL/JSON v1.0.6';
  End Get_Version;

End Json_Parser;
/

Create Or Replace Package Json_Printer As
  Indent_String Varchar2(10 Char) := '  '; --chr(9); for tab
  Newline_Char  Varchar2(2 Char) := Chr(13) || Chr(10); -- Windows style
  --newline_char varchar2(2) := chr(10); -- Mac style
  --newline_char varchar2(2) := chr(13); -- Linux style
  Ascii_Output   Boolean Not Null := True;
  Escape_Solidus Boolean Not Null := False;

  Function Pretty_Print(Obj Json, Spaces Boolean Default True, Line_Length Number Default 0) Return Varchar2;
  Function Pretty_Print_List(Obj Json_List, Spaces Boolean Default True, Line_Length Number Default 0) Return Varchar2;
  Function Pretty_Print_Any(Json_Part Json_Value, Spaces Boolean Default True, Line_Length Number Default 0)
    Return Varchar2;
  Procedure Pretty_Print(Obj Json, Spaces Boolean Default True, Buf In Out Nocopy Clob, Line_Length Number Default 0, Erase_Clob Boolean Default True);
  Procedure Pretty_Print_List(Obj Json_List, Spaces Boolean Default True, Buf In Out Nocopy Clob, Line_Length Number Default 0, Erase_Clob Boolean Default True);
  Procedure Pretty_Print_Any(Json_Part Json_Value, Spaces Boolean Default True, Buf In Out Nocopy Clob, Line_Length Number Default 0, Erase_Clob Boolean Default True);

  Procedure Dbms_Output_Clob(My_Clob Clob, Delim Varchar2, Jsonp Varchar2 Default Null);
  Procedure Htp_Output_Clob(My_Clob Clob, Jsonp Varchar2 Default Null);
  -- made public just for testing/profiling... 
  Function Escapestring(Str Varchar2) Return Varchar2;

End Json_Printer;
/


Create Or Replace Package Body Json_Printer As
  Max_Line_Len Number := 0;
  Cur_Line_Len Number := 0;
  Type Tmap_Char_String Is Table Of Varchar2(40) Index By Varchar2(1 Char); /* index by unicode char */
  Char_Map                Tmap_Char_String;
  Char_Map_Escape_Solidus Boolean := Escape_Solidus;
  Char_Map_Ascii_Output   Boolean := Ascii_Output;

  Function Llcheck(Str In Varchar2) Return Varchar2 As
  Begin
    If (Max_Line_Len > 0 And Length(Str) + Cur_Line_Len > Max_Line_Len) Then
      Cur_Line_Len := Length(Str);
      Return Newline_Char || Str;
    Else
      Cur_Line_Len := Cur_Line_Len + Length(Str);
      Return Str;
    End If;
  End Llcheck;

  -- escapes a single character.
  Function Escapechar(Ch Char) Return Varchar2 Deterministic Is
    Result Varchar2(20);
  Begin
    --backspace b = U+0008
    --formfeed  f = U+000C
    --newline   n = U+000A
    --carret    r = U+000D
    --tabulator t = U+0009
    Result := Ch;
  
    Case Ch
      When Chr(8) Then
        Result := '\b';
      When Chr(9) Then
        Result := '\t';
      When Chr(10) Then
        Result := '\n';
      When Chr(12) Then
        Result := '\f';
      When Chr(13) Then
        Result := '\r';
      When Chr(34) Then
        Result := '\"';
      When Chr(47) Then
        If (Escape_Solidus) Then
          Result := '\/';
        End If;
      When Chr(92) Then
        Result := '\\';
      Else
        If (Ascii(Ch) < 32) Then
          Result := '\u' || Replace(Substr(To_Char(Ascii(Ch), 'XXXX'), 2, 4), ' ', '0');
        Elsif (Ascii_Output) Then
          Result := Replace(Asciistr(Ch), '\', '\u');
        End If;
    End Case;
    Return Result;
  End;

  Function Escapestring(Str Varchar2) Return Varchar2 As
    Sb  Varchar2(32767 Byte) := '';
    Buf Varchar2(40);
    Ch  Varchar2(1 Char); /* unicode char */
  Begin
    If (Str Is Null) Then
      Return '';
    End If;
  
    -- clear the cache if global parameters have been changed
    If Char_Map_Escape_Solidus <> Escape_Solidus Or Char_Map_Ascii_Output <> Ascii_Output Then
      Char_Map.Delete;
      Char_Map_Escape_Solidus := Escape_Solidus;
      Char_Map_Ascii_Output   := Ascii_Output;
    End If;
  
    For i In 1 .. Length(Str) Loop
      Ch := Substr(Str, i, 1);
    
      Begin
        -- it this char has already been processed, I have cached its escaped value
        Buf := Char_Map(Ch);
      Exception
        When No_Data_Found Then
          -- otherwise, i convert the value and add it to the cache
          Buf := Escapechar(Ch);
          Char_Map(Ch) := Buf;
      End;
    
      Sb := Sb || Buf;
    End Loop;
    Return Sb;
  End Escapestring;

  Function Newline(Spaces Boolean) Return Varchar2 As
  Begin
    Cur_Line_Len := 0;
    If (Spaces) Then
      Return Newline_Char;
    Else
      Return '';
    End If;
  End;

  Function Tab(Indent Number, Spaces Boolean) Return Varchar2 As
    i Varchar(200) := '';
  Begin
    If (Not Spaces) Then
      Return '';
    End If;
    For x In 1 .. Indent Loop
      i := i || Indent_String;
    End Loop;
    Return i;
  End;

  Function Getcommasep(Spaces Boolean) Return Varchar2 As
  Begin
    If (Spaces) Then
      Return ', ';
    Else
      Return ',';
    End If;
  End;

  Function Getmemname(Mem Json_Value, Spaces Boolean) Return Varchar2 As
  Begin
    If (Spaces) Then
      Return Llcheck('"' || Escapestring(Mem.Mapname) || '"') || Llcheck(' : ');
    Else
      Return Llcheck('"' || Escapestring(Mem.Mapname) || '"') || Llcheck(':');
    End If;
  End;

  /* Clob method start here */
  Procedure Add_To_Clob(Buf_Lob In Out Nocopy Clob, Buf_Str In Out Nocopy Varchar2, Str Varchar2) As
  Begin
    If (Lengthb(Str) > 32767 - Lengthb(Buf_Str)) Then
      --      dbms_lob.append(buf_lob, buf_str);
      Dbms_Lob.Writeappend(Buf_Lob, Length(Buf_Str), Buf_Str);
      Buf_Str := Str;
    Else
      Buf_Str := Buf_Str || Str;
    End If;
  End Add_To_Clob;

  Procedure Flush_Clob(Buf_Lob In Out Nocopy Clob, Buf_Str In Out Nocopy Varchar2) As
  Begin
    --    dbms_lob.append(buf_lob, buf_str);
    Dbms_Lob.Writeappend(Buf_Lob, Length(Buf_Str), Buf_Str);
  End Flush_Clob;

  Procedure Ppobj(Obj Json, Indent Number, Buf In Out Nocopy Clob, Spaces Boolean, Buf_Str In Out Nocopy Varchar2);

  Procedure Ppstring(Elem Json_Value, Buf In Out Nocopy Clob, Buf_Str In Out Nocopy Varchar2) Is
    Offset Number := 1;
    v_Str  Varchar(5000 Char);
    Amount Number := 5000; /*chunk size for use in escapeString. Maximum escaped unicode string size for chunk may be 6 one-byte chars * 5000 chunk size in multi-byte chars = 30000 1-byte chars. Maximum value may be 32767 1-byte chars */
  Begin
    Add_To_Clob(Buf, Buf_Str, Case When Elem.Num = 1 Then '"' Else '/**/' End);
    If (Elem.Extended_Str Is Not Null) Then
      --clob implementation
      While (Offset <= Dbms_Lob.Getlength(Elem.Extended_Str)) Loop
        Dbms_Lob.Read(Elem.Extended_Str, Amount, Offset, v_Str);
        If (Elem.Num = 1) Then
          Add_To_Clob(Buf, Buf_Str, Escapestring(v_Str));
        Else
          Add_To_Clob(Buf, Buf_Str, v_Str);
        End If;
        Offset := Offset + Amount;
      End Loop;
    Else
      If (Elem.Num = 1) Then
        While (Offset <= Length(Elem.Str)) Loop
          v_Str := Substr(Elem.Str, Offset, Amount);
          Add_To_Clob(Buf, Buf_Str, Escapestring(v_Str));
          Offset := Offset + Amount;
        End Loop;
      Else
        Add_To_Clob(Buf, Buf_Str, Elem.Str);
      End If;
    End If;
    Add_To_Clob(Buf, Buf_Str, Case When Elem.Num = 1 Then '"' Else '/**/' End || Newline_Char);
  End;

  Procedure Ppea(Input Json_List, Indent Number, Buf In Out Nocopy Clob, Spaces Boolean, Buf_Str In Out Nocopy Varchar2) As
    Elem   Json_Value;
    Arr    Json_Value_Array := Input.List_Data;
    Numbuf Varchar2(4000);
  Begin
    For y In 1 .. Arr.Count Loop
      Elem := Arr(y);
      If (Elem Is Not Null) Then
        Case Elem.Get_Type
          When 'number' Then
            Numbuf := '';
            If (Elem.Get_Number < 1 And Elem.Get_Number > 0) Then
              Numbuf := '0';
            End If;
            If (Elem.Get_Number < 0 And Elem.Get_Number > -1) Then
              Numbuf := '-0';
              Numbuf := Numbuf || Substr(To_Char(Elem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,'''), 2);
            Else
              Numbuf := Numbuf || To_Char(Elem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
            End If;
            Add_To_Clob(Buf, Buf_Str, Llcheck(Numbuf));
          When 'string' Then
            Ppstring(Elem, Buf, Buf_Str);
          When 'bool' Then
            If (Elem.Get_Bool) Then
              Add_To_Clob(Buf, Buf_Str, Llcheck('true'));
            Else
              Add_To_Clob(Buf, Buf_Str, Llcheck('false'));
            End If;
          When 'null' Then
            Add_To_Clob(Buf, Buf_Str, Llcheck('null'));
          When 'array' Then
            Add_To_Clob(Buf, Buf_Str, Llcheck('['));
            Ppea(Json_List(Elem), Indent, Buf, Spaces, Buf_Str);
            Add_To_Clob(Buf, Buf_Str, Llcheck(']'));
          When 'object' Then
            Ppobj(Json(Elem), Indent, Buf, Spaces, Buf_Str);
          Else
            Add_To_Clob(Buf, Buf_Str, Llcheck(Elem.Get_Type));
        End Case;
      End If;
      If (y != Arr.Count) Then
        Add_To_Clob(Buf, Buf_Str, Llcheck(Getcommasep(Spaces)));
      End If;
    End Loop;
  End Ppea;

  Procedure Ppmem(Mem Json_Value, Indent Number, Buf In Out Nocopy Clob, Spaces Boolean, Buf_Str In Out Nocopy Varchar2) As
    Numbuf Varchar2(4000);
  Begin
    Add_To_Clob(Buf, Buf_Str, Llcheck(Tab(Indent, Spaces)) || Llcheck(Getmemname(Mem, Spaces)));
    Case Mem.Get_Type
      When 'number' Then
        If (Mem.Get_Number < 1 And Mem.Get_Number > 0) Then
          Numbuf := '0';
        End If;
        If (Mem.Get_Number < 0 And Mem.Get_Number > -1) Then
          Numbuf := '-0';
          Numbuf := Numbuf || Substr(To_Char(Mem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,'''), 2);
        Else
          Numbuf := Numbuf || To_Char(Mem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
        End If;
        Add_To_Clob(Buf, Buf_Str, Llcheck(Numbuf));
      When 'string' Then
        Ppstring(Mem, Buf, Buf_Str);
      When 'bool' Then
        If (Mem.Get_Bool) Then
          Add_To_Clob(Buf, Buf_Str, Llcheck('true'));
        Else
          Add_To_Clob(Buf, Buf_Str, Llcheck('false'));
        End If;
      When 'null' Then
        Add_To_Clob(Buf, Buf_Str, Llcheck('null'));
      When 'array' Then
        Add_To_Clob(Buf, Buf_Str, Llcheck('['));
        Ppea(Json_List(Mem), Indent, Buf, Spaces, Buf_Str);
        Add_To_Clob(Buf, Buf_Str, Llcheck(']'));
      When 'object' Then
        Ppobj(Json(Mem), Indent, Buf, Spaces, Buf_Str);
      Else
        Add_To_Clob(Buf, Buf_Str, Llcheck(Mem.Get_Type));
    End Case;
  End Ppmem;

  Procedure Ppobj(Obj Json, Indent Number, Buf In Out Nocopy Clob, Spaces Boolean, Buf_Str In Out Nocopy Varchar2) As
  Begin
    Add_To_Clob(Buf, Buf_Str, Llcheck('{') || Newline(Spaces));
    For m In 1 .. Obj.Json_Data.Count Loop
      Ppmem(Obj.Json_Data(m), Indent + 1, Buf, Spaces, Buf_Str);
      If (m != Obj.Json_Data.Count) Then
        Add_To_Clob(Buf, Buf_Str, Llcheck(',') || Newline(Spaces));
      Else
        Add_To_Clob(Buf, Buf_Str, Newline(Spaces));
      End If;
    End Loop;
    Add_To_Clob(Buf, Buf_Str, Llcheck(Tab(Indent, Spaces)) || Llcheck('}')); -- || chr(13);
  End Ppobj;

  Procedure Pretty_Print(Obj Json, Spaces Boolean Default True, Buf In Out Nocopy Clob, Line_Length Number Default 0, Erase_Clob Boolean Default True) As
    Buf_Str Varchar2(32767);
    Amount  Number := Dbms_Lob.Getlength(Buf);
  Begin
    If (Erase_Clob And Amount > 0) Then
      Dbms_Lob.Trim(Buf, 0);
      Dbms_Lob.Erase(Buf, Amount);
    End If;
  
    Max_Line_Len := Line_Length;
    Cur_Line_Len := 0;
    Ppobj(Obj, 0, Buf, Spaces, Buf_Str);
    Flush_Clob(Buf, Buf_Str);
  End;

  Procedure Pretty_Print_List(Obj Json_List, Spaces Boolean Default True, Buf In Out Nocopy Clob, Line_Length Number Default 0, Erase_Clob Boolean Default True) As
    Buf_Str Varchar2(32767);
    Amount  Number := Dbms_Lob.Getlength(Buf);
  Begin
    If (Erase_Clob And Amount > 0) Then
      Dbms_Lob.Trim(Buf, 0);
      Dbms_Lob.Erase(Buf, Amount);
    End If;
  
    Max_Line_Len := Line_Length;
    Cur_Line_Len := 0;
    Add_To_Clob(Buf, Buf_Str, Llcheck('['));
    Ppea(Obj, 0, Buf, Spaces, Buf_Str);
    Add_To_Clob(Buf, Buf_Str, Llcheck(']'));
    Flush_Clob(Buf, Buf_Str);
  End;

  Procedure Pretty_Print_Any(Json_Part Json_Value, Spaces Boolean Default True, Buf In Out Nocopy Clob, Line_Length Number Default 0, Erase_Clob Boolean Default True) As
    Buf_Str Varchar2(32767) := '';
    Numbuf  Varchar2(4000);
    Amount  Number := Dbms_Lob.Getlength(Buf);
  Begin
    If (Erase_Clob And Amount > 0) Then
      Dbms_Lob.Trim(Buf, 0);
      Dbms_Lob.Erase(Buf, Amount);
    End If;
  
    Case Json_Part.Get_Type
      When 'number' Then
        If (Json_Part.Get_Number < 1 And Json_Part.Get_Number > 0) Then
          Numbuf := '0';
        End If;
        If (Json_Part.Get_Number < 0 And Json_Part.Get_Number > -1) Then
          Numbuf := '-0';
          Numbuf := Numbuf || Substr(To_Char(Json_Part.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,'''), 2);
        Else
          Numbuf := Numbuf || To_Char(Json_Part.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
        End If;
        Add_To_Clob(Buf, Buf_Str, Numbuf);
      When 'string' Then
        Ppstring(Json_Part, Buf, Buf_Str);
      When 'bool' Then
        If (Json_Part.Get_Bool) Then
          Add_To_Clob(Buf, Buf_Str, 'true');
        Else
          Add_To_Clob(Buf, Buf_Str, 'false');
        End If;
      When 'null' Then
        Add_To_Clob(Buf, Buf_Str, 'null');
      When 'array' Then
        Pretty_Print_List(Json_List(Json_Part), Spaces, Buf, Line_Length);
        Return;
      When 'object' Then
        Pretty_Print(Json(Json_Part), Spaces, Buf, Line_Length);
        Return;
      Else
        Add_To_Clob(Buf, Buf_Str, 'unknown type:' || Json_Part.Get_Type);
    End Case;
    Flush_Clob(Buf, Buf_Str);
  End;

  /* Clob method end here */

  /* Varchar2 method start here */
  Procedure Add_Buf(Buf In Out Nocopy Varchar2, Str In Varchar2) As
  Begin
    If (Lengthb(Str) > 32767 - Lengthb(Buf)) Then
      Raise_Application_Error(-20001, 'Length of result JSON more than 32767 bytes. Use to_clob() procedures');
    End If;
    Buf := Buf || Str;
  End;

  Procedure Ppstring(Elem Json_Value, Buf In Out Nocopy Varchar2) Is
    Offset Number := 1;
    v_Str  Varchar(5000 Char);
    Amount Number := 5000; /*chunk size for use in escapeString. Maximum escaped unicode string size for chunk may be 6 one-byte chars * 5000 chunk size in multi-byte chars = 30000 1-byte chars. Maximum value may be 32767 1-byte chars */
  Begin
    Add_Buf(Buf, Case When Elem.Num = 1 Then '"' Else '/**/' End);
    If (Elem.Extended_Str Is Not Null) Then
      --clob implementation
      While (Offset <= Dbms_Lob.Getlength(Elem.Extended_Str)) Loop
        Dbms_Lob.Read(Elem.Extended_Str, Amount, Offset, v_Str);
        If (Elem.Num = 1) Then
          Add_Buf(Buf, Escapestring(v_Str));
        Else
          Add_Buf(Buf, v_Str);
        End If;
        Offset := Offset + Amount;
      End Loop;
    Else
      If (Elem.Num = 1) Then
        While (Offset <= Length(Elem.Str)) Loop
          v_Str := Substr(Elem.Str, Offset, Amount);
          Add_Buf(Buf, Escapestring(v_Str));
          Offset := Offset + Amount;
        End Loop;
      Else
        Add_Buf(Buf, Elem.Str);
      End If;
    End If;
    Add_Buf(Buf, Case When Elem.Num = 1 Then '"' Else '/**/' End || Newline_Char);
  End;

  Procedure Ppobj(Obj Json, Indent Number, Buf In Out Nocopy Varchar2, Spaces Boolean);

  Procedure Ppea(Input Json_List, Indent Number, Buf In Out Varchar2, Spaces Boolean) As
    Elem Json_Value;
    Arr  Json_Value_Array := Input.List_Data;
    Str  Varchar2(400);
  Begin
    For y In 1 .. Arr.Count Loop
      Elem := Arr(y);
      If (Elem Is Not Null) Then
        Case Elem.Get_Type
          When 'number' Then
            Str := '';
            If (Elem.Get_Number < 1 And Elem.Get_Number > 0) Then
              Str := '0';
            End If;
            If (Elem.Get_Number < 0 And Elem.Get_Number > -1) Then
              Str := '-0' || Substr(To_Char(Elem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,'''), 2);
            Else
              Str := Str || To_Char(Elem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
            End If;
            Add_Buf(Buf, Llcheck(Str));
          When 'string' Then
            Ppstring(Elem, Buf);
          When 'bool' Then
            If (Elem.Get_Bool) Then
              Add_Buf(Buf, Llcheck('true'));
            Else
              Add_Buf(Buf, Llcheck('false'));
            End If;
          When 'null' Then
            Add_Buf(Buf, Llcheck('null'));
          When 'array' Then
            Add_Buf(Buf, Llcheck('['));
            Ppea(Json_List(Elem), Indent, Buf, Spaces);
            Add_Buf(Buf, Llcheck(']'));
          When 'object' Then
            Ppobj(Json(Elem), Indent, Buf, Spaces);
          Else
            Add_Buf(Buf, Llcheck(Elem.Get_Type)); /* should never happen */
        End Case;
      End If;
      If (y != Arr.Count) Then
        Add_Buf(Buf, Llcheck(Getcommasep(Spaces)));
      End If;
    End Loop;
  End Ppea;

  Procedure Ppmem(Mem Json_Value, Indent Number, Buf In Out Nocopy Varchar2, Spaces Boolean) As
    Str Varchar2(400) := '';
  Begin
    Add_Buf(Buf, Llcheck(Tab(Indent, Spaces)) || Getmemname(Mem, Spaces));
    Case Mem.Get_Type
      When 'number' Then
        If (Mem.Get_Number < 1 And Mem.Get_Number > 0) Then
          Str := '0';
        End If;
        If (Mem.Get_Number < 0 And Mem.Get_Number > -1) Then
          Str := '-0' || Substr(To_Char(Mem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,'''), 2);
        Else
          Str := Str || To_Char(Mem.Get_Number, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
        End If;
        Add_Buf(Buf, Llcheck(Str));
      When 'string' Then
        Ppstring(Mem, Buf);
      When 'bool' Then
        If (Mem.Get_Bool) Then
          Add_Buf(Buf, Llcheck('true'));
        Else
          Add_Buf(Buf, Llcheck('false'));
        End If;
      When 'null' Then
        Add_Buf(Buf, Llcheck('null'));
      When 'array' Then
        Add_Buf(Buf, Llcheck('['));
        Ppea(Json_List(Mem), Indent, Buf, Spaces);
        Add_Buf(Buf, Llcheck(']'));
      When 'object' Then
        Ppobj(Json(Mem), Indent, Buf, Spaces);
      Else
        Add_Buf(Buf, Llcheck(Mem.Get_Type)); /* should never happen */
    End Case;
  End Ppmem;

  Procedure Ppobj(Obj Json, Indent Number, Buf In Out Nocopy Varchar2, Spaces Boolean) As
  Begin
    Add_Buf(Buf, Llcheck('{') || Newline(Spaces));
    For m In 1 .. Obj.Json_Data.Count Loop
      Ppmem(Obj.Json_Data(m), Indent + 1, Buf, Spaces);
      If (m != Obj.Json_Data.Count) Then
        Add_Buf(Buf, Llcheck(',') || Newline(Spaces));
      Else
        Add_Buf(Buf, Newline(Spaces));
      End If;
    End Loop;
    Add_Buf(Buf, Llcheck(Tab(Indent, Spaces)) || Llcheck('}')); -- || chr(13);
  End Ppobj;

  Function Pretty_Print(Obj Json, Spaces Boolean Default True, Line_Length Number Default 0) Return Varchar2 As
    Buf Varchar2(32767 Byte) := '';
  Begin
    Max_Line_Len := Line_Length;
    Cur_Line_Len := 0;
    Ppobj(Obj, 0, Buf, Spaces);
    Return Buf;
  End Pretty_Print;

  Function Pretty_Print_List(Obj Json_List, Spaces Boolean Default True, Line_Length Number Default 0) Return Varchar2 As
    Buf Varchar2(32767 Byte) := '';
  Begin
    Max_Line_Len := Line_Length;
    Cur_Line_Len := 0;
    Add_Buf(Buf, Llcheck('['));
    Ppea(Obj, 0, Buf, Spaces);
    Add_Buf(Buf, Llcheck(']'));
    Return Buf;
  End;

  Function Pretty_Print_Any(Json_Part Json_Value, Spaces Boolean Default True, Line_Length Number Default 0)
    Return Varchar2 As
    Buf Varchar2(32767) := '';
  Begin
    Case Json_Part.Get_Type
      When 'number' Then
        If (Json_Part.Get_Number() < 1 And Json_Part.Get_Number() > 0) Then
          Buf := '0';
        End If;
        If (Json_Part.Get_Number() < 0 And Json_Part.Get_Number() > -1) Then
          Buf := '-0';
          Buf := Buf || Substr(To_Char(Json_Part.Get_Number(), 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,'''), 2);
        Else
          Buf := Buf || To_Char(Json_Part.Get_Number(), 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
        End If;
      When 'string' Then
        Ppstring(Json_Part, Buf);
      When 'bool' Then
        If (Json_Part.Get_Bool) Then
          Buf := 'true';
        Else
          Buf := 'false';
        End If;
      When 'null' Then
        Buf := 'null';
      When 'array' Then
        Buf := Pretty_Print_List(Json_List(Json_Part), Spaces, Line_Length);
      When 'object' Then
        Buf := Pretty_Print(Json(Json_Part), Spaces, Line_Length);
      Else
        Buf := 'weird error: ' || Json_Part.Get_Type;
    End Case;
    Return Buf;
  End;

  Procedure Dbms_Output_Clob(My_Clob Clob, Delim Varchar2, Jsonp Varchar2 Default Null) As
    Prev       Number := 1;
    Indx       Number := 1;
    Size_Of_Nl Number := Lengthb(Delim);
    v_Str      Varchar2(32767);
    Amount     Number := 8191; /* max unicode chars */
  Begin
    If (Jsonp Is Not Null) Then
      Dbms_Output.Put_Line(Jsonp || '(');
    End If;
    While (Indx != 0) Loop
      --read every line
      Indx := Dbms_Lob.Instr(My_Clob, Delim, Prev + 1);
      --     dbms_output.put_line(prev || ' to ' || indx);
    
      If (Indx = 0) Then
        --emit from prev to end;
        Amount := 8191; /* max unicode chars */
        Loop
          Dbms_Lob.Read(My_Clob, Amount, Prev, v_Str);
          Dbms_Output.Put_Line(v_Str);
          Prev := Prev + Amount - 1;
          Exit When Prev >= Dbms_Lob.Getlength(My_Clob);
        End Loop;
      Else
        Amount := Indx - Prev;
        If (Amount > 8191) Then
          /* max unicode chars */
          Amount := 8191; /* max unicode chars */
          Loop
            Dbms_Lob.Read(My_Clob, Amount, Prev, v_Str);
            Dbms_Output.Put_Line(v_Str);
            Prev   := Prev + Amount - 1;
            Amount := Indx - Prev;
            Exit When Prev >= Indx - 1;
            If (Amount > 8191) Then
              Amount := 8191;
            End If; /* max unicode chars */
          End Loop;
          Prev := Indx + Size_Of_Nl;
        Else
          Dbms_Lob.Read(My_Clob, Amount, Prev, v_Str);
          Dbms_Output.Put_Line(v_Str);
          Prev := Indx + Size_Of_Nl;
        End If;
      End If;
    
    End Loop;
    If (Jsonp Is Not Null) Then
      Dbms_Output.Put_Line(')');
    End If;
  End;

  Procedure Htp_Output_Clob(My_Clob Clob, Jsonp Varchar2 Default Null) As
    l_Amt Number Default 30;
    l_Off Number Default 1;
    l_Str Varchar2(4096);
  
  Begin
    If (Jsonp Is Not Null) Then
      Htp.Prn(Jsonp || '(');
    End If;
  
    Begin
      Loop
        Dbms_Lob.Read(My_Clob, l_Amt, l_Off, l_Str);
      
        -- it is vital to use htp.PRN to avoid
        -- spurious line feeds getting added to your
        -- document
        Htp.Prn(l_Str);
        l_Off := l_Off + l_Amt;
        l_Amt := 4096;
      End Loop;
    Exception
      When No_Data_Found Then
        Null;
    End;
  
    If (Jsonp Is Not Null) Then
      Htp.Prn(')');
    End If;
  End;

End Json_Printer;
/


Create Or Replace Type Body Json_Value As

  Constructor Function Json_Value(Object_Or_Array Sys.Anydata) Return Self As Result As
  Begin
    Case Object_Or_Array.Gettypename
      When Sys_Context('userenv', 'current_schema') || '.JSON_LIST' Then
        Self.Typeval := 2;
      When Sys_Context('userenv', 'current_schema') || '.JSON' Then
        Self.Typeval := 1;
      Else
        Raise_Application_Error(-20102, 'JSON_Value init error (JSON or JSON\_List allowed)');
    End Case;
    Self.Object_Or_Array := Object_Or_Array;
    If (Self.Object_Or_Array Is Null) Then
      Self.Typeval := 6;
    End If;
  
    Return;
  End Json_Value;

  Constructor Function Json_Value(Str Varchar2, Esc Boolean Default True) Return Self As Result As
  Begin
    Self.Typeval := 3;
    If (Esc) Then
      Self.Num := 1;
    Else
      Self.Num := 0;
    End If; --message to pretty printer
    Self.Str := Str;
    Return;
  End Json_Value;

  Constructor Function Json_Value(Str Clob, Esc Boolean Default True) Return Self As Result As
    Amount Number := 5000; /* for Unicode text, varchar2 self.str not exceed 5000 chars, does not limit size of data */
  Begin
    Self.Typeval := 3;
    If (Esc) Then
      Self.Num := 1;
    Else
      Self.Num := 0;
    End If; --message to pretty printer
    If (Dbms_Lob.Getlength(Str) > Amount) Then
      Extended_Str := Str;
    End If;
    -- GHS 20120615: Added IF structure to handle null clobs
    If Dbms_Lob.Getlength(Str) > 0 Then
      Dbms_Lob.Read(Str, Amount, 1, Self.Str);
    End If;
    Return;
  End Json_Value;

  Constructor Function Json_Value(Num Number) Return Self As Result As
  Begin
    Self.Typeval := 4;
    Self.Num     := Num;
    If (Self.Num Is Null) Then
      Self.Typeval := 6;
    End If;
    Return;
  End Json_Value;

  Constructor Function Json_Value(b Boolean) Return Self As Result As
  Begin
    Self.Typeval := 5;
    Self.Num     := 0;
    If (b) Then
      Self.Num := 1;
    End If;
    If (b Is Null) Then
      Self.Typeval := 6;
    End If;
    Return;
  End Json_Value;

  Constructor Function Json_Value Return Self As Result As
  Begin
    Self.Typeval := 6; /* for JSON null */
    Return;
  End Json_Value;

  Static Function Makenull Return Json_Value As
  Begin
    Return Json_Value;
  End Makenull;

  Member Function Get_Type Return Varchar2 As
  Begin
    Case Self.Typeval
      When 1 Then
        Return 'object';
      When 2 Then
        Return 'array';
      When 3 Then
        Return 'string';
      When 4 Then
        Return 'number';
      When 5 Then
        Return 'bool';
      When 6 Then
        Return 'null';
    End Case;
  
    Return 'unknown type';
  End Get_Type;

  Member Function Get_String(Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null) Return Varchar2 As
  Begin
    If (Self.Typeval = 3) Then
      If (Max_Byte_Size Is Not Null) Then
        Return Substrb(Self.Str, 1, Max_Byte_Size);
      Elsif (Max_Char_Size Is Not Null) Then
        Return Substr(Self.Str, 1, Max_Char_Size);
      Else
        Return Self.Str;
      End If;
    End If;
    Return Null;
  End Get_String;

  Member Procedure Get_String(Self In Json_Value, Buf In Out Nocopy Clob) As
  Begin
    If (Self.Typeval = 3) Then
      If (Extended_Str Is Not Null) Then
        Dbms_Lob.Copy(Buf, Extended_Str, Dbms_Lob.Getlength(Extended_Str));
      Else
        Dbms_Lob.Writeappend(Buf, Length(Self.Str), Self.Str);
      End If;
    End If;
  End Get_String;

  Member Function Get_Number Return Number As
  Begin
    If (Self.Typeval = 4) Then
      Return Self.Num;
    End If;
    Return Null;
  End Get_Number;

  Member Function Get_Bool Return Boolean As
  Begin
    If (Self.Typeval = 5) Then
      Return Self.Num = 1;
    End If;
    Return Null;
  End Get_Bool;

  Member Function Get_Null Return Varchar2 As
  Begin
    If (Self.Typeval = 6) Then
      Return 'null';
    End If;
    Return Null;
  End Get_Null;

  Member Function Is_Object Return Boolean As
  Begin
    Return Self.Typeval = 1;
  End;
  Member Function Is_Array Return Boolean As
  Begin
    Return Self.Typeval = 2;
  End;
  Member Function Is_String Return Boolean As
  Begin
    Return Self.Typeval = 3;
  End;
  Member Function Is_Number Return Boolean As
  Begin
    Return Self.Typeval = 4;
  End;
  Member Function Is_Bool Return Boolean As
  Begin
    Return Self.Typeval = 5;
  End;
  Member Function Is_Null Return Boolean As
  Begin
    Return Self.Typeval = 6;
  End;

  /* Output methods */
  Member Function To_Char(Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2 As
  Begin
    If (Spaces Is Null) Then
      Return Json_Printer.Pretty_Print_Any(Self, Line_Length => Chars_Per_Line);
    Else
      Return Json_Printer.Pretty_Print_Any(Self, Spaces, Line_Length => Chars_Per_Line);
    End If;
  End;

  Member Procedure To_Clob(Self In Json_Value, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True) As
  Begin
    If (Spaces Is Null) Then
      Json_Printer.Pretty_Print_Any(Self, False, Buf, Line_Length => Chars_Per_Line, Erase_Clob => Erase_Clob);
    Else
      Json_Printer.Pretty_Print_Any(Self, Spaces, Buf, Line_Length => Chars_Per_Line, Erase_Clob => Erase_Clob);
    End If;
  End;

  Member Procedure Print(Self In Json_Value, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null) As
    --32512 is the real maximum in sqldeveloper
    My_Clob Clob;
  Begin
    My_Clob := Empty_Clob();
    Dbms_Lob.Createtemporary(My_Clob, True);
    Json_Printer.Pretty_Print_Any(Self, Spaces, My_Clob, Case When(Chars_Per_Line > 32512) Then 32512 Else
                                   Chars_Per_Line End);
    Json_Printer.Dbms_Output_Clob(My_Clob, Json_Printer.Newline_Char, Jsonp);
    Dbms_Lob.Freetemporary(My_Clob);
  End;

  Member Procedure Htp(Self In Json_Value, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null) As
    My_Clob Clob;
  Begin
    My_Clob := Empty_Clob();
    Dbms_Lob.Createtemporary(My_Clob, True);
    Json_Printer.Pretty_Print_Any(Self, Spaces, My_Clob, Chars_Per_Line);
    Json_Printer.Htp_Output_Clob(My_Clob, Jsonp);
    Dbms_Lob.Freetemporary(My_Clob);
  End;

  Member Function Value_Of(Self In Json_Value, Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null)
    Return Varchar2 As
  Begin
    Case Self.Typeval
      When 1 Then
        Return 'json object';
      When 2 Then
        Return 'json array';
      When 3 Then
        Return Self.Get_String(Max_Byte_Size, Max_Char_Size);
      When 4 Then
        Return Self.Get_Number();
      When 5 Then
        If (Self.Get_Bool()) Then
          Return 'true';
        Else
          Return 'false';
        End If;
      Else
        Return Null;
    End Case;
  End;

End;
/

Create Or Replace Package Json_Ext As
  /* This package contains extra methods to lookup types and
  an easy way of adding date values in json - without changing the structure */
  Function Parsepath(Json_Path Varchar2, Base Number Default 1) Return Json_List;

  --JSON Path getters
  Function Get_Json_Value(Obj Json, v_Path Varchar2, Base Number Default 1) Return Json_Value;
  Function Get_String(Obj Json, Path Varchar2, Base Number Default 1) Return Varchar2;
  Function Get_Number(Obj Json, Path Varchar2, Base Number Default 1) Return Number;
  Function Get_Json(Obj Json, Path Varchar2, Base Number Default 1) Return Json;
  Function Get_Json_List(Obj Json, Path Varchar2, Base Number Default 1) Return Json_List;
  Function Get_Bool(Obj Json, Path Varchar2, Base Number Default 1) Return Boolean;

  --JSON Path putters
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Varchar2, Base Number Default 1);
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Number, Base Number Default 1);
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Json, Base Number Default 1);
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Json_List, Base Number Default 1);
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Boolean, Base Number Default 1);
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Json_Value, Base Number Default 1);

  Procedure Remove(Obj In Out Nocopy Json, Path Varchar2, Base Number Default 1);

  --Pretty print with JSON Path - obsolete in 0.9.4 - obj.path(v_path).(to_char,print,htp)
  Function Pp(Obj Json, v_Path Varchar2) Return Varchar2;
  Procedure Pp(Obj Json, v_Path Varchar2); --using dbms_output.put_line
  Procedure Pp_Htp(Obj Json, v_Path Varchar2); --using htp.print

  --extra function checks if number has no fraction
  Function Is_Integer(v Json_Value) Return Boolean;

  Format_String Varchar2(30 Char) := 'yyyy-mm-dd hh24:mi:ss';
  --extension enables json to store dates without comprimising the implementation
  Function To_Json_Value(d Date) Return Json_Value;
  --notice that a date type in json is also a varchar2
  Function Is_Date(v Json_Value) Return Boolean;
  --convertion is needed to extract dates 
  Function To_Date(v Json_Value) Return Date;
  -- alias so that old code doesnt break
  Function To_Date2(v Json_Value) Return Date;
  --JSON Path with date
  Function Get_Date(Obj Json, Path Varchar2, Base Number Default 1) Return Date;
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Date, Base Number Default 1);

  --experimental support of binary data with base64
  Function Base64(Binarydata Blob) Return Json_List;
  Function Base64(l Json_List) Return Blob;

  Function Encode(Binarydata Blob) Return Json_Value;
  Function Decode(v Json_Value) Return Blob;

End Json_Ext;
/


Create Or Replace Package Body Json_Ext As
  Scanner_Exception Exception;
  Pragma Exception_Init(Scanner_Exception, -20100);
  Parser_Exception Exception;
  Pragma Exception_Init(Parser_Exception, -20101);
  Jext_Exception Exception;
  Pragma Exception_Init(Jext_Exception, -20110);

  --extra function checks if number has no fraction
  Function Is_Integer(v Json_Value) Return Boolean As
    Myint Number(38); --the oracle way to specify an integer
  Begin
    If (v.Is_Number) Then
      Myint := v.Get_Number;
      Return(Myint = v.Get_Number); --no rounding errors
    Else
      Return False;
    End If;
  End;

  --extension enables json to store dates without comprimising the implementation
  Function To_Json_Value(d Date) Return Json_Value As
  Begin
    Return Json_Value(To_Char(d, Format_String));
  End;

  --notice that a date type in json is also a varchar2
  Function Is_Date(v Json_Value) Return Boolean As
    Temp Date;
  Begin
    Temp := Json_Ext.To_Date(v);
    Return True;
  Exception
    When Others Then
      Return False;
  End;

  --convertion is needed to extract dates
  Function To_Date(v Json_Value) Return Date As
  Begin
    If (v.Is_String) Then
      Return Standard.To_Date(v.Get_String, Format_String);
    Else
      Raise_Application_Error(-20110, 'Anydata did not contain a date-value');
    End If;
  Exception
    When Others Then
      Raise_Application_Error(-20110, 'Anydata did not contain a date on the format: ' || Format_String);
  End;

  -- alias so that old code doesn t break
  Function To_Date2(v Json_Value) Return Date As
  Begin
    Return To_Date(v);
  End;

  --Json Path parser
  Function Parsepath(Json_Path Varchar2, Base Number Default 1) Return Json_List As
    Build_Path Varchar2(32767) := '[';
    Buf        Varchar2(4);
    Endstring  Varchar2(1);
    Indx       Number := 1;
    Ret        Json_List;
  
    Procedure Next_Char As
    Begin
      If (Indx <= Length(Json_Path)) Then
        Buf  := Substr(Json_Path, Indx, 1);
        Indx := Indx + 1;
      Else
        Buf := Null;
      End If;
    End;
    --skip ws
    Procedure Skipws As
    Begin
      While (Buf In (Chr(9), Chr(10), Chr(13), ' ')) Loop
        Next_Char;
      End Loop;
    End;
  
  Begin
    Next_Char();
    While (Buf Is Not Null) Loop
      If (Buf = '.') Then
        Next_Char();
        If (Buf Is Null) Then
          Raise_Application_Error(-20110, 'JSON Path parse error: . is not a valid json_path end');
        End If;
        If (Not Regexp_Like(Buf, '^[[:alnum:]\_ ]+', 'c')) Then
          Raise_Application_Error(-20110, 'JSON Path parse error: alpha-numeric character or space expected at position ' || Indx);
        End If;
      
        If (Build_Path != '[') Then
          Build_Path := Build_Path || ',';
        End If;
        Build_Path := Build_Path || '"';
        While (Regexp_Like(Buf, '^[[:alnum:]\_ ]+', 'c')) Loop
          Build_Path := Build_Path || Buf;
          Next_Char();
        End Loop;
        Build_Path := Build_Path || '"';
      Elsif (Buf = '[') Then
        Next_Char();
        Skipws();
        If (Buf Is Null) Then
          Raise_Application_Error(-20110, 'JSON Path parse error: [ is not a valid json_path end');
        End If;
        If (Buf In ('1', '2', '3', '4', '5', '6', '7', '8', '9') Or (Buf = '0' And Base = 0)) Then
          If (Build_Path != '[') Then
            Build_Path := Build_Path || ',';
          End If;
          While (Buf In ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9')) Loop
            Build_Path := Build_Path || Buf;
            Next_Char();
          End Loop;
        Elsif (Regexp_Like(Buf, '^(\"|\'')', 'c')) Then
          Endstring := Buf;
          If (Build_Path != '[') Then
            Build_Path := Build_Path || ',';
          End If;
          Build_Path := Build_Path || '"';
          Next_Char();
          If (Buf Is Null) Then
            Raise_Application_Error(-20110, 'JSON Path parse error: premature json_path end');
          End If;
          While (Buf != Endstring) Loop
            Build_Path := Build_Path || Buf;
            Next_Char();
            If (Buf Is Null) Then
              Raise_Application_Error(-20110, 'JSON Path parse error: premature json_path end');
            End If;
            If (Buf = '\') Then
              Next_Char();
              Build_Path := Build_Path || '\' || Buf;
              Next_Char();
            End If;
          End Loop;
          Build_Path := Build_Path || '"';
          Next_Char();
        Else
          Raise_Application_Error(-20110, 'JSON Path parse error: expected a string or an positive integer at ' || Indx);
        End If;
        Skipws();
        If (Buf Is Null) Then
          Raise_Application_Error(-20110, 'JSON Path parse error: premature json_path end');
        End If;
        If (Buf != ']') Then
          Raise_Application_Error(-20110, 'JSON Path parse error: no array ending found. found: ' || Buf);
        End If;
        Next_Char();
        Skipws();
      Elsif (Build_Path = '[') Then
        If (Not Regexp_Like(Buf, '^[[:alnum:]\_ ]+', 'c')) Then
          Raise_Application_Error(-20110, 'JSON Path parse error: alpha-numeric character or space expected at position ' || Indx);
        End If;
        Build_Path := Build_Path || '"';
        While (Regexp_Like(Buf, '^[[:alnum:]\_ ]+', 'c')) Loop
          Build_Path := Build_Path || Buf;
          Next_Char();
        End Loop;
        Build_Path := Build_Path || '"';
      Else
        Raise_Application_Error(-20110, 'JSON Path parse error: expected . or [ found ' || Buf || ' at position ' || Indx);
      End If;
    
    End Loop;
  
    Build_Path := Build_Path || ']';
    Build_Path := Replace(Replace(Replace(Replace(Replace(Build_Path, Chr(9), '\t'), Chr(10), '\n'), Chr(13), '\f'), Chr(8), '\b'), Chr(14), '\r');
  
    Ret := Json_List(Build_Path);
    If (Base != 1) Then
      --fix base 0 to base 1
      Declare
        Elem Json_Value;
      Begin
        For i In 1 .. Ret.Count Loop
          Elem := Ret.Get(i);
          If (Elem.Is_Number) Then
            Ret.Replace(i, Elem.Get_Number() + 1);
          End If;
        End Loop;
      End;
    End If;
  
    Return Ret;
  End Parsepath;

  --JSON Path getters
  Function Get_Json_Value(Obj Json, v_Path Varchar2, Base Number Default 1) Return Json_Value As
    Path Json_List;
    Ret  Json_Value;
    o    Json;
    l    Json_List;
  Begin
    Path := Parsepath(v_Path, Base);
    Ret  := Obj.To_Json_Value;
    If (Path.Count = 0) Then
      Return Ret;
    End If;
  
    For i In 1 .. Path.Count Loop
      If (Path.Get(i).Is_String()) Then
        --string fetch only on json
        o   := Json(Ret);
        Ret := o.Get(Path.Get(i).Get_String());
      Else
        --number fetch on json and json_list
        If (Ret.Is_Array()) Then
          l   := Json_List(Ret);
          Ret := l.Get(Path.Get(i).Get_Number());
        Else
          o   := Json(Ret);
          l   := o.Get_Values();
          Ret := l.Get(Path.Get(i).Get_Number());
        End If;
      End If;
    End Loop;
  
    Return Ret;
  Exception
    When Scanner_Exception Then
      Raise;
    When Parser_Exception Then
      Raise;
    When Jext_Exception Then
      Raise;
    When Others Then
      Return Null;
  End Get_Json_Value;

  --JSON Path getters
  Function Get_String(Obj Json, Path Varchar2, Base Number Default 1) Return Varchar2 As
    Temp Json_Value;
  Begin
    Temp := Get_Json_Value(Obj, Path, Base);
    If (Temp Is Null Or Not Temp.Is_String) Then
      Return Null;
    Else
      Return Temp.Get_String;
    End If;
  End;

  Function Get_Number(Obj Json, Path Varchar2, Base Number Default 1) Return Number As
    Temp Json_Value;
  Begin
    Temp := Get_Json_Value(Obj, Path, Base);
    If (Temp Is Null Or Not Temp.Is_Number) Then
      Return Null;
    Else
      Return Temp.Get_Number;
    End If;
  End;

  Function Get_Json(Obj Json, Path Varchar2, Base Number Default 1) Return Json As
    Temp Json_Value;
  Begin
    Temp := Get_Json_Value(Obj, Path, Base);
    If (Temp Is Null Or Not Temp.Is_Object) Then
      Return Null;
    Else
      Return Json(Temp);
    End If;
  End;

  Function Get_Json_List(Obj Json, Path Varchar2, Base Number Default 1) Return Json_List As
    Temp Json_Value;
  Begin
    Temp := Get_Json_Value(Obj, Path, Base);
    If (Temp Is Null Or Not Temp.Is_Array) Then
      Return Null;
    Else
      Return Json_List(Temp);
    End If;
  End;

  Function Get_Bool(Obj Json, Path Varchar2, Base Number Default 1) Return Boolean As
    Temp Json_Value;
  Begin
    Temp := Get_Json_Value(Obj, Path, Base);
    If (Temp Is Null Or Not Temp.Is_Bool) Then
      Return Null;
    Else
      Return Temp.Get_Bool;
    End If;
  End;

  Function Get_Date(Obj Json, Path Varchar2, Base Number Default 1) Return Date As
    Temp Json_Value;
  Begin
    Temp := Get_Json_Value(Obj, Path, Base);
    If (Temp Is Null Or Not Is_Date(Temp)) Then
      Return Null;
    Else
      Return Json_Ext.To_Date(Temp);
    End If;
  End;

  /* JSON Path putter internal function */
  Procedure Put_Internal(Obj In Out Nocopy Json, v_Path Varchar2, Elem Json_Value, Base Number) As
    Val           Json_Value := Elem;
    Path          Json_List;
    Backreference Json_List := Json_List();
  
    Keyval    Json_Value;
    Keynum    Number;
    Keystring Varchar2(4000);
    Temp      Json_Value := Obj.To_Json_Value;
    Obj_Temp  Json;
    List_Temp Json_List;
    Inserter  Json_Value;
  Begin
    Path := Json_Ext.Parsepath(v_Path, Base);
    If (Path.Count = 0) Then
      Raise_Application_Error(-20110, 'JSON_EXT put error: cannot put with empty string.');
    End If;
  
    --build backreference
    For i In 1 .. Path.Count Loop
      --backreference.print(false);
      Keyval := Path.Get(i);
      If (Keyval.Is_Number()) Then
        --nummer index
        Keynum := Keyval.Get_Number();
        If ((Not Temp.Is_Object()) And (Not Temp.Is_Array())) Then
          If (Val Is Null) Then
            Return;
          End If;
          Backreference.Remove_Last;
          Temp := Json_List().To_Json_Value();
          Backreference.Append(Temp);
        End If;
      
        If (Temp.Is_Object()) Then
          Obj_Temp := Json(Temp);
          If (Obj_Temp.Count < Keynum) Then
            If (Val Is Null) Then
              Return;
            End If;
            Raise_Application_Error(-20110, 'JSON_EXT put error: access object with to few members.');
          End If;
          Temp := Obj_Temp.Get(Keynum);
        Else
          List_Temp := Json_List(Temp);
          If (List_Temp.Count < Keynum) Then
            If (Val Is Null) Then
              Return;
            End If;
            --raise error or quit if val is null
            For i In List_Temp.Count + 1 .. Keynum Loop
              List_Temp.Append(Json_Value.Makenull);
            End Loop;
            Backreference.Remove_Last;
            Backreference.Append(List_Temp);
          End If;
        
          Temp := List_Temp.Get(Keynum);
        End If;
      Else
        --streng index
        Keystring := Keyval.Get_String();
        If (Not Temp.Is_Object()) Then
          --backreference.print;
          If (Val Is Null) Then
            Return;
          End If;
          Backreference.Remove_Last;
          Temp := Json().To_Json_Value();
          Backreference.Append(Temp);
          --raise_application_error(-20110, JSON_ext put error: trying to access a non object with a string.); 
        End If;
        Obj_Temp := Json(Temp);
        Temp     := Obj_Temp.Get(Keystring);
      End If;
    
      If (Temp Is Null) Then
        If (Val Is Null) Then
          Return;
        End If;
        --what to expect
        Keyval := Path.Get(i + 1);
        If (Keyval Is Not Null And Keyval.Is_Number()) Then
          Temp := Json_List().To_Json_Value;
        Else
          Temp := Json().To_Json_Value;
        End If;
      End If;
      Backreference.Append(Temp);
    End Loop;
  
    --  backreference.print(false);
    --  path.print(false);
  
    --use backreference and path together
    Inserter := Val;
    For i In Reverse 1 .. Backreference.Count Loop
      --    inserter.print(false);
      If (i = 1) Then
        Keyval := Path.Get(1);
        If (Keyval.Is_String()) Then
          Keystring := Keyval.Get_String();
        Else
          Keynum := Keyval.Get_Number();
          Declare
            T1 Json_Value := Obj.Get(Keynum);
          Begin
            Keystring := T1.Mapname;
          End;
        End If;
        If (Inserter Is Null) Then
          Obj.Remove(Keystring);
        Else
          Obj.Put(Keystring, Inserter);
        End If;
      Else
        Temp := Backreference.Get(i - 1);
        If (Temp.Is_Object()) Then
          Keyval   := Path.Get(i);
          Obj_Temp := Json(Temp);
          If (Keyval.Is_String()) Then
            Keystring := Keyval.Get_String();
          Else
            Keynum := Keyval.Get_Number();
            Declare
              T1 Json_Value := Obj_Temp.Get(Keynum);
            Begin
              Keystring := T1.Mapname;
            End;
          End If;
          If (Inserter Is Null) Then
            Obj_Temp.Remove(Keystring);
            If (Obj_Temp.Count > 0) Then
              Inserter := Obj_Temp.To_Json_Value;
            End If;
          Else
            Obj_Temp.Put(Keystring, Inserter);
            Inserter := Obj_Temp.To_Json_Value;
          End If;
        Else
          --array only number
          Keynum    := Path.Get(i).Get_Number();
          List_Temp := Json_List(Temp);
          List_Temp.Remove(Keynum);
          If (Not Inserter Is Null) Then
            List_Temp.Append(Inserter, Keynum);
            Inserter := List_Temp.To_Json_Value;
          Else
            If (List_Temp.Count > 0) Then
              Inserter := List_Temp.To_Json_Value;
            End If;
          End If;
        End If;
      End If;
    
    End Loop;
  
  End Put_Internal;

  /* JSON Path putters */
  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Varchar2, Base Number Default 1) As
  Begin
    If Elem Is Null Then
      Put_Internal(Obj, Path, Json_Value(), Base);
    Else
      Put_Internal(Obj, Path, Json_Value(Elem), Base);
    End If;
  End;

  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Number, Base Number Default 1) As
  Begin
    If Elem Is Null Then
      Put_Internal(Obj, Path, Json_Value(), Base);
    Else
      Put_Internal(Obj, Path, Json_Value(Elem), Base);
    End If;
  End;

  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Json, Base Number Default 1) As
  Begin
    If Elem Is Null Then
      Put_Internal(Obj, Path, Json_Value(), Base);
    Else
      Put_Internal(Obj, Path, Elem.To_Json_Value, Base);
    End If;
  End;

  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Json_List, Base Number Default 1) As
  Begin
    If Elem Is Null Then
      Put_Internal(Obj, Path, Json_Value(), Base);
    Else
      Put_Internal(Obj, Path, Elem.To_Json_Value, Base);
    End If;
  End;

  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Boolean, Base Number Default 1) As
  Begin
    If Elem Is Null Then
      Put_Internal(Obj, Path, Json_Value(), Base);
    Else
      Put_Internal(Obj, Path, Json_Value(Elem), Base);
    End If;
  End;

  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Json_Value, Base Number Default 1) As
  Begin
    If Elem Is Null Then
      Put_Internal(Obj, Path, Json_Value(), Base);
    Else
      Put_Internal(Obj, Path, Elem, Base);
    End If;
  End;

  Procedure Put(Obj In Out Nocopy Json, Path Varchar2, Elem Date, Base Number Default 1) As
  Begin
    If Elem Is Null Then
      Put_Internal(Obj, Path, Json_Value(), Base);
    Else
      Put_Internal(Obj, Path, Json_Ext.To_Json_Value(Elem), Base);
    End If;
  End;

  Procedure Remove(Obj In Out Nocopy Json, Path Varchar2, Base Number Default 1) As
  Begin
    Json_Ext.Put_Internal(Obj, Path, Null, Base);
    --    if(json_ext.get_json_value(obj,path) is not null) then
    --    end if;
  End Remove;

  --Pretty print with JSON Path
  Function Pp(Obj Json, v_Path Varchar2) Return Varchar2 As
    Json_Part Json_Value;
  Begin
    Json_Part := Json_Ext.Get_Json_Value(Obj, v_Path);
    If (Json_Part Is Null) Then
      Return '';
    Else
      Return Json_Printer.Pretty_Print_Any(Json_Part); --escapes a possible internal string
    End If;
  End Pp;

  Procedure Pp(Obj Json, v_Path Varchar2) As --using dbms_output.put_line
  Begin
    Dbms_Output.Put_Line(Pp(Obj, v_Path));
  End Pp;

  -- spaces = false!
  Procedure Pp_Htp(Obj Json, v_Path Varchar2) As
    --using htp.print
    Json_Part Json_Value;
  Begin
    Json_Part := Json_Ext.Get_Json_Value(Obj, v_Path);
    If (Json_Part Is Null) Then
      Htp.Print;
    Else
      Htp.Print(Json_Printer.Pretty_Print_Any(Json_Part, False));
    End If;
  End Pp_Htp;

  Function Base64(Binarydata Blob) Return Json_List As
    Obj  Json_List := Json_List();
    c    Clob := Empty_Clob();
    Benc Blob;
  
    v_Blob_Offset  Number := 1;
    v_Clob_Offset  Number := 1;
    v_Lang_Context Number := Dbms_Lob.Default_Lang_Ctx;
    v_Warning      Number;
    v_Amount       Pls_Integer;
    --    temp varchar2(32767);
  
    Function Encodeblob2base64(Pblobin In Blob) Return Blob Is
      Vamount     Number := 45;
      Vblobenc    Blob := Empty_Blob();
      Vblobenclen Number := 0;
      Vblobinlen  Number := 0;
      Vbuffer     Raw(45);
      Voffset     Number := 1;
    Begin
      --      dbms_output.put_line('Start base64 encoding.');
      Vblobinlen := Dbms_Lob.Getlength(Pblobin);
      --      dbms_output.put_line('<BlobInLength>' || vBlobInLen);
      Dbms_Lob.Createtemporary(Vblobenc, True);
      Loop
        If Voffset >= Vblobinlen Then
          Exit;
        End If;
        Dbms_Lob.Read(Pblobin, Vamount, Voffset, Vbuffer);
        Begin
          Dbms_Lob.Append(Vblobenc, Utl_Encode.Base64_Encode(Vbuffer));
        Exception
          When Others Then
            Dbms_Output.Put_Line('<vAmount>' || Vamount || '<vOffset>' || Voffset || '<vBuffer>' || Vbuffer);
            Dbms_Output.Put_Line('ERROR IN append: ' || Sqlerrm);
            Raise;
        End;
        Voffset := Voffset + Vamount;
      End Loop;
      Vblobenclen := Dbms_Lob.Getlength(Vblobenc);
      --      dbms_output.put_line(<BlobEncLength> || vBlobEncLen);
      --      dbms_output.put_line(Finshed base64 encoding.);
      Return Vblobenc;
    End Encodeblob2base64;
  Begin
    Benc := Encodeblob2base64(Binarydata);
    Dbms_Lob.Createtemporary(c, True);
    v_Amount := Dbms_Lob.Getlength(Benc);
    Dbms_Lob.Converttoclob(c, Benc, v_Amount, v_Clob_Offset, v_Blob_Offset, 1, v_Lang_Context, v_Warning);
  
    v_Amount      := Dbms_Lob.Getlength(c);
    v_Clob_Offset := 1;
    --dbms_output.put_line(V amount: ||v_amount);
    While (v_Clob_Offset < v_Amount) Loop
      --dbms_output.put_line(v_offset);
      --temp := ;
      --dbms_output.put_line(size: ||length(temp));
      Obj.Append(Dbms_Lob.Substr(c, 4000, v_Clob_Offset));
      v_Clob_Offset := v_Clob_Offset + 4000;
    End Loop;
    Dbms_Lob.Freetemporary(Benc);
    Dbms_Lob.Freetemporary(c);
    --dbms_output.put_line(obj.count);
    --dbms_output.put_line(obj.get_last().to_char);
    Return Obj;
  
  End Base64;

  Function Base64(l Json_List) Return Blob As
    c    Clob := Empty_Clob();
    b    Blob := Empty_Blob();
    Bret Blob;
  
    v_Blob_Offset  Number := 1;
    v_Clob_Offset  Number := 1;
    v_Lang_Context Number := 0; --DBMS_LOB.DEFAULT_LANG_CTX;
    v_Warning      Number;
    v_Amount       Pls_Integer;
  
    Function Decodebase642blob(Pblobin In Blob) Return Blob Is
      Vamount     Number := 256; --32;
      Vblobdec    Blob := Empty_Blob();
      Vblobdeclen Number := 0;
      Vblobinlen  Number := 0;
      Vbuffer     Raw(256); --32);
      Voffset     Number := 1;
    Begin
      --      dbms_output.put_line(Start base64 decoding.);
      Vblobinlen := Dbms_Lob.Getlength(Pblobin);
      --      dbms_output.put_line(<BlobInLength> || vBlobInLen);
      Dbms_Lob.Createtemporary(Vblobdec, True);
      Loop
        If Voffset >= Vblobinlen Then
          Exit;
        End If;
        Dbms_Lob.Read(Pblobin, Vamount, Voffset, Vbuffer);
        Begin
          Dbms_Lob.Append(Vblobdec, Utl_Encode.Base64_Decode(Vbuffer));
        Exception
          When Others Then
            Dbms_Output.Put_Line('<vAmount>' || Vamount || '<vOffset>' || Voffset || '<vBuffer>' || Vbuffer);
            Dbms_Output.Put_Line('ERROR IN append: ' || Sqlerrm);
            Raise;
        End;
        Voffset := Voffset + Vamount;
      End Loop;
      Vblobdeclen := Dbms_Lob.Getlength(Vblobdec);
      --      dbms_output.put_line(<BlobDecLength> || vBlobDecLen);
      --      dbms_output.put_line('Finshed base64 decoding.');
      Return Vblobdec;
    End Decodebase642blob;
  Begin
    Dbms_Lob.Createtemporary(c, True);
    For i In 1 .. l.Count Loop
      Dbms_Lob.Append(c, l.Get(i).Get_String());
    End Loop;
    v_Amount := Dbms_Lob.Getlength(c);
    --    dbms_output.put_line(L C||v_amount);
  
    Dbms_Lob.Createtemporary(b, True);
    Dbms_Lob.Converttoblob(b, c, Dbms_Lob.Lobmaxsize, v_Clob_Offset, v_Blob_Offset, 1, v_Lang_Context, v_Warning);
    Dbms_Lob.Freetemporary(c);
    v_Amount := Dbms_Lob.Getlength(b);
    --    dbms_output.put_line(L B||v_amount);
  
    Bret := Decodebase642blob(b);
    Dbms_Lob.Freetemporary(b);
    Return Bret;
  
  End Base64;

  Function Encode(Binarydata Blob) Return Json_Value As
    Obj  Json_Value;
    c    Clob := Empty_Clob();
    Benc Blob;
  
    v_Blob_Offset  Number := 1;
    v_Clob_Offset  Number := 1;
    v_Lang_Context Number := Dbms_Lob.Default_Lang_Ctx;
    v_Warning      Number;
    v_Amount       Pls_Integer;
    --    temp varchar2(32767);
  
    Function Encodeblob2base64(Pblobin In Blob) Return Blob Is
      Vamount     Number := 45;
      Vblobenc    Blob := Empty_Blob();
      Vblobenclen Number := 0;
      Vblobinlen  Number := 0;
      Vbuffer     Raw(45);
      Voffset     Number := 1;
    Begin
      --      dbms_output.put_line('Start base64 encoding.');
      Vblobinlen := Dbms_Lob.Getlength(Pblobin);
      --      dbms_output.put_line('<BlobInLength>' || vBlobInLen);
      Dbms_Lob.Createtemporary(Vblobenc, True);
      Loop
        If Voffset >= Vblobinlen Then
          Exit;
        End If;
        Dbms_Lob.Read(Pblobin, Vamount, Voffset, Vbuffer);
        Begin
          Dbms_Lob.Append(Vblobenc, Utl_Encode.Base64_Encode(Vbuffer));
        Exception
          When Others Then
            Dbms_Output.Put_Line('<vAmount>' || Vamount || '<vOffset>' || Voffset || '<vBuffer>' || Vbuffer);
            Dbms_Output.Put_Line('ERROR IN append: ' || Sqlerrm);
            Raise;
        End;
        Voffset := Voffset + Vamount;
      End Loop;
      Vblobenclen := Dbms_Lob.Getlength(Vblobenc);
      --      dbms_output.put_line('<BlobEncLength>' || vBlobEncLen);
      --      dbms_output.put_line('Finshed base64 encoding.');
      Return Vblobenc;
    End Encodeblob2base64;
  Begin
    Benc := Encodeblob2base64(Binarydata);
    Dbms_Lob.Createtemporary(c, True);
    v_Amount := Dbms_Lob.Getlength(Benc);
    Dbms_Lob.Converttoclob(c, Benc, v_Amount, v_Clob_Offset, v_Blob_Offset, 1, v_Lang_Context, v_Warning);
  
    Obj := Json_Value(c);
  
    Dbms_Lob.Freetemporary(Benc);
    Dbms_Lob.Freetemporary(c);
    --dbms_output.put_line(obj.count);
    --dbms_output.put_line(obj.get_last().to_char);
    Return Obj;
  
  End Encode;

  Function Decode(v Json_Value) Return Blob As
    c    Clob := Empty_Clob();
    b    Blob := Empty_Blob();
    Bret Blob;
  
    v_Blob_Offset  Number := 1;
    v_Clob_Offset  Number := 1;
    v_Lang_Context Number := 0; --DBMS_LOB.DEFAULT_LANG_CTX;
    v_Warning      Number;
    v_Amount       Pls_Integer;
  
    Function Decodebase642blob(Pblobin In Blob) Return Blob Is
      Vamount     Number := 256; --32;
      Vblobdec    Blob := Empty_Blob();
      Vblobdeclen Number := 0;
      Vblobinlen  Number := 0;
      Vbuffer     Raw(256); --32);
      Voffset     Number := 1;
    Begin
      --      dbms_output.put_line('Start base64 decoding.');
      Vblobinlen := Dbms_Lob.Getlength(Pblobin);
      --      dbms_output.put_line('<BlobInLength>' || vBlobInLen);
      Dbms_Lob.Createtemporary(Vblobdec, True);
      Loop
        If Voffset >= Vblobinlen Then
          Exit;
        End If;
        Dbms_Lob.Read(Pblobin, Vamount, Voffset, Vbuffer);
        Begin
          Dbms_Lob.Append(Vblobdec, Utl_Encode.Base64_Decode(Vbuffer));
        Exception
          When Others Then
            Dbms_Output.Put_Line('<vAmount>' || Vamount || '<vOffset>' || Voffset || '<vBuffer>' || Vbuffer);
            Dbms_Output.Put_Line('ERROR IN append: ' || Sqlerrm);
            Raise;
        End;
        Voffset := Voffset + Vamount;
      End Loop;
      Vblobdeclen := Dbms_Lob.Getlength(Vblobdec);
      --      dbms_output.put_line('<BlobDecLength>' || vBlobDecLen);
      --      dbms_output.put_line('Finshed base64 decoding.');
      Return Vblobdec;
    End Decodebase642blob;
  Begin
    Dbms_Lob.Createtemporary(c, True);
    v.Get_String(c);
    v_Amount := Dbms_Lob.Getlength(c);
    --    dbms_output.put_line('L C'||v_amount);
  
    Dbms_Lob.Createtemporary(b, True);
    Dbms_Lob.Converttoblob(b, c, Dbms_Lob.Lobmaxsize, v_Clob_Offset, v_Blob_Offset, 1, v_Lang_Context, v_Warning);
    Dbms_Lob.Freetemporary(c);
    v_Amount := Dbms_Lob.Getlength(b);
    --    dbms_output.put_line('L B'||v_amount);
  
    Bret := Decodebase642blob(b);
    Dbms_Lob.Freetemporary(b);
    Return Bret;
  
  End Decode;

End Json_Ext;
/

Create Or Replace Type Body Json As

  /* Constructors */
  Constructor Function Json Return Self As Result As
  Begin
    Self.Json_Data           := Json_Value_Array();
    Self.Check_For_Duplicate := 1;
    Return;
  End;

  Constructor Function Json(Str Varchar2) Return Self As Result As
  Begin
    Self                     := Json_Parser.Parser(Str);
    Self.Check_For_Duplicate := 1;
    Return;
  End;

  Constructor Function Json(Str In Clob) Return Self As Result As
  Begin
    Self                     := Json_Parser.Parser(Str);
    Self.Check_For_Duplicate := 1;
    Return;
  End;

  Constructor Function Json(Cast Json_Value) Return Self As Result As
    x Number;
  Begin
    x                        := Cast.Object_Or_Array.Getobject(Self);
    Self.Check_For_Duplicate := 1;
    Return;
  End;

  Constructor Function Json(l In Out Nocopy Json_List) Return Self As Result As
  Begin
    For i In 1 .. l.List_Data.Count Loop
      If (l.List_Data(i).Mapname Is Null Or l.List_Data(i).Mapname Like 'row%') Then
        l.List_Data(i).Mapname := 'row' || i;
      End If;
      l.List_Data(i).Mapindx := i;
    End Loop;
  
    Self.Json_Data           := l.List_Data;
    Self.Check_For_Duplicate := 1;
    Return;
  End;

  /* Member setter methods */
  Member Procedure Remove(Self In Out Nocopy Json, Pair_Name Varchar2) As
    Temp Json_Value;
    Indx Pls_Integer;
  
    Function Get_Member(Pair_Name Varchar2) Return Json_Value As
      Indx Pls_Integer;
    Begin
      Indx := Json_Data.First;
      Loop
        Exit When Indx Is Null;
        If (Pair_Name Is Null And Json_Data(Indx).Mapname Is Null) Then
          Return Json_Data(Indx);
        End If;
        If (Json_Data(Indx).Mapname = Pair_Name) Then
          Return Json_Data(Indx);
        End If;
        Indx := Json_Data.Next(Indx);
      End Loop;
      Return Null;
    End;
  Begin
    Temp := Get_Member(Pair_Name);
    If (Temp Is Null) Then
      Return;
    End If;
  
    Indx := Json_Data.Next(Temp.Mapindx);
    Loop
      Exit When Indx Is Null;
      Json_Data(Indx).Mapindx := Indx - 1;
      Json_Data(Indx - 1) := Json_Data(Indx);
      Indx := Json_Data.Next(Indx);
    End Loop;
    Json_Data.Trim(1);
    --num_elements := num_elements - 1;
  End;

  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_Value, Position Pls_Integer Default Null) As
    Insert_Value Json_Value := Nvl(Pair_Value, Json_Value.Makenull);
    Indx         Pls_Integer;
    x            Number;
    Temp         Json_Value;
    Function Get_Member(Pair_Name Varchar2) Return Json_Value As
      Indx Pls_Integer;
    Begin
      Indx := Json_Data.First;
      Loop
        Exit When Indx Is Null;
        If (Pair_Name Is Null And Json_Data(Indx).Mapname Is Null) Then
          Return Json_Data(Indx);
        End If;
        If (Json_Data(Indx).Mapname = Pair_Name) Then
          Return Json_Data(Indx);
        End If;
        Indx := Json_Data.Next(Indx);
      End Loop;
      Return Null;
    End;
  Begin
    Insert_Value.Mapname := Pair_Name;
    --    self.remove(pair_name);
    If (Self.Check_For_Duplicate = 1) Then
      Temp := Get_Member(Pair_Name);
    Else
      Temp := Null;
    End If;
    If (Temp Is Not Null) Then
      Insert_Value.Mapindx := Temp.Mapindx;
      Json_Data(Temp.Mapindx) := Insert_Value;
      Return;
    Elsif (Position Is Null Or Position > Self.Count) Then
      Json_Data.Extend(1);
      Json_Data(Json_Data.Count) := Insert_Value;
      Json_Data(Json_Data.Count).Mapindx := Json_Data.Count;
    Elsif (Position < 2) Then
      --insert at the start of the list
      Indx := Json_Data.Last;
      Json_Data.Extend;
      Loop
        Exit When Indx Is Null;
        Temp := Json_Data(Indx);
        Temp.Mapindx := Indx + 1;
        Json_Data(Temp.Mapindx) := Temp;
        Indx := Json_Data.Prior(Indx);
      End Loop;
      Json_Data(1) := Insert_Value;
      Insert_Value.Mapindx := 1;
    Else
      --insert somewhere in the list
      Indx := Json_Data.Last;
      --      dbms_output.put_line('Test '||indx);
      Json_Data.Extend;
      --      dbms_output.put_line('Test '||indx);
      Loop
        --        dbms_output.put_line('Test '||indx);
        Temp := Json_Data(Indx);
        Temp.Mapindx := Indx + 1;
        Json_Data(Temp.Mapindx) := Temp;
        Exit When Indx = Position;
        Indx := Json_Data.Prior(Indx);
      End Loop;
      Json_Data(Position) := Insert_Value;
      Json_Data(Position).Mapindx := Position;
    End If;
    --    num_elements := num_elements + 1;
  End;

  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Varchar2, Position Pls_Integer Default Null) As
  Begin
    Put(Pair_Name, Json_Value(Pair_Value), Position);
  End;

  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Number, Position Pls_Integer Default Null) As
  Begin
    If (Pair_Value Is Null) Then
      Put(Pair_Name, Json_Value(), Position);
    Else
      Put(Pair_Name, Json_Value(Pair_Value), Position);
    End If;
  End;

  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Boolean, Position Pls_Integer Default Null) As
  Begin
    If (Pair_Value Is Null) Then
      Put(Pair_Name, Json_Value(), Position);
    Else
      Put(Pair_Name, Json_Value(Pair_Value), Position);
    End If;
  End;

  Member Procedure Check_Duplicate(Self In Out Nocopy Json, v_Set Boolean) As
  Begin
    If (v_Set) Then
      Check_For_Duplicate := 1;
    Else
      Check_For_Duplicate := 0;
    End If;
  End;

  /* deprecated putters */

  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json, Position Pls_Integer Default Null) As
  Begin
    If (Pair_Value Is Null) Then
      Put(Pair_Name, Json_Value(), Position);
    Else
      Put(Pair_Name, Pair_Value.To_Json_Value, Position);
    End If;
  End;

  Member Procedure Put(Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_List, Position Pls_Integer Default Null) As
  Begin
    If (Pair_Value Is Null) Then
      Put(Pair_Name, Json_Value(), Position);
    Else
      Put(Pair_Name, Pair_Value.To_Json_Value, Position);
    End If;
  End;

  /* Member getter methods */
  Member Function Count Return Number As
  Begin
    Return Self.Json_Data.Count;
  End;

  Member Function Get(Pair_Name Varchar2) Return Json_Value As
    Indx Pls_Integer;
  Begin
    Indx := Json_Data.First;
    Loop
      Exit When Indx Is Null;
      If (Pair_Name Is Null And Json_Data(Indx).Mapname Is Null) Then
        Return Json_Data(Indx);
      End If;
      If (Json_Data(Indx).Mapname = Pair_Name) Then
        Return Json_Data(Indx);
      End If;
      Indx := Json_Data.Next(Indx);
    End Loop;
    Return Null;
  End;

  Member Function Get(Position Pls_Integer) Return Json_Value As
  Begin
    If (Self.Count >= Position And Position > 0) Then
      Return Self.Json_Data(Position);
    End If;
    Return Null; -- do not throw error, just return null
  End;

  Member Function Index_Of(Pair_Name Varchar2) Return Number As
    Indx Pls_Integer;
  Begin
    Indx := Json_Data.First;
    Loop
      Exit When Indx Is Null;
      If (Pair_Name Is Null And Json_Data(Indx).Mapname Is Null) Then
        Return Indx;
      End If;
      If (Json_Data(Indx).Mapname = Pair_Name) Then
        Return Indx;
      End If;
      Indx := Json_Data.Next(Indx);
    End Loop;
    Return - 1;
  End;

  Member Function Exist(Pair_Name Varchar2) Return Boolean As
  Begin
    Return(Self.Get(Pair_Name) Is Not Null);
  End;

  /* Output methods */
  Member Function To_Char(Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2 As
  Begin
    If (Spaces Is Null) Then
      Return Json_Printer.Pretty_Print(Self, Line_Length => Chars_Per_Line);
    Else
      Return Json_Printer.Pretty_Print(Self, Spaces, Line_Length => Chars_Per_Line);
    End If;
  End;

  Member Procedure To_Clob(Self In Json, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True) As
  Begin
    If (Spaces Is Null) Then
      Json_Printer.Pretty_Print(Self, False, Buf, Line_Length => Chars_Per_Line, Erase_Clob => Erase_Clob);
    Else
      Json_Printer.Pretty_Print(Self, Spaces, Buf, Line_Length => Chars_Per_Line, Erase_Clob => Erase_Clob);
    End If;
  End;

  Member Procedure Print(Self In Json, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null) As
    --32512 is the real maximum in sqldeveloper
    My_Clob Clob;
  Begin
    My_Clob := Empty_Clob();
    Dbms_Lob.Createtemporary(My_Clob, True);
    Json_Printer.Pretty_Print(Self, Spaces, My_Clob, Case When(Chars_Per_Line > 32512) Then 32512 Else Chars_Per_Line End);
    Json_Printer.Dbms_Output_Clob(My_Clob, Json_Printer.Newline_Char, Jsonp);
    Dbms_Lob.Freetemporary(My_Clob);
  End;

  Member Procedure Htp(Self In Json, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null) As
    My_Clob Clob;
  Begin
    My_Clob := Empty_Clob();
    Dbms_Lob.Createtemporary(My_Clob, True);
    Json_Printer.Pretty_Print(Self, Spaces, My_Clob, Chars_Per_Line);
    Json_Printer.Htp_Output_Clob(My_Clob, Jsonp);
    Dbms_Lob.Freetemporary(My_Clob);
  End;

  Member Function To_Json_Value Return Json_Value As
  Begin
    Return Json_Value(Sys.Anydata.Convertobject(Self));
  End;

  /* json path */
  Member Function Path(Json_Path Varchar2, Base Number Default 1) Return Json_Value As
  Begin
    Return Json_Ext.Get_Json_Value(Self, Json_Path, Base);
  End Path;

  /* json path_put */
  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_Value, Base Number Default 1) As
  Begin
    Json_Ext.Put(Self, Json_Path, Elem, Base);
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Varchar2, Base Number Default 1) As
  Begin
    Json_Ext.Put(Self, Json_Path, Elem, Base);
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Number, Base Number Default 1) As
  Begin
    If (Elem Is Null) Then
      Json_Ext.Put(Self, Json_Path, Json_Value(), Base);
    Else
      Json_Ext.Put(Self, Json_Path, Elem, Base);
    End If;
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Boolean, Base Number Default 1) As
  Begin
    If (Elem Is Null) Then
      Json_Ext.Put(Self, Json_Path, Json_Value(), Base);
    Else
      Json_Ext.Put(Self, Json_Path, Elem, Base);
    End If;
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_List, Base Number Default 1) As
  Begin
    If (Elem Is Null) Then
      Json_Ext.Put(Self, Json_Path, Json_Value(), Base);
    Else
      Json_Ext.Put(Self, Json_Path, Elem, Base);
    End If;
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json, Json_Path Varchar2, Elem Json, Base Number Default 1) As
  Begin
    If (Elem Is Null) Then
      Json_Ext.Put(Self, Json_Path, Json_Value(), Base);
    Else
      Json_Ext.Put(Self, Json_Path, Elem, Base);
    End If;
  End Path_Put;

  Member Procedure Path_Remove(Self In Out Nocopy Json, Json_Path Varchar2, Base Number Default 1) As
  Begin
    Json_Ext.Remove(Self, Json_Path, Base);
  End Path_Remove;

  /* Thanks to Matt Nolan */
  Member Function Get_Keys Return Json_List As
    Keys Json_List;
    Indx Pls_Integer;
  Begin
    Keys := Json_List();
    Indx := Json_Data.First;
    Loop
      Exit When Indx Is Null;
      Keys.Append(Json_Data(Indx).Mapname);
      Indx := Json_Data.Next(Indx);
    End Loop;
    Return Keys;
  End;

  Member Function Get_Values Return Json_List As
    Vals Json_List := Json_List();
  Begin
    Vals.List_Data := Self.Json_Data;
    Return Vals;
  End;

  Member Procedure Remove_Duplicates(Self In Out Nocopy Json) As
  Begin
    Json_Parser.Remove_Duplicates(Self);
  End Remove_Duplicates;

End;
/

Create Or Replace Type Body Json_List As

  Constructor Function Json_List Return Self As Result As
  Begin
    Self.List_Data := Json_Value_Array();
    Return;
  End;

  Constructor Function Json_List(Str Varchar2) Return Self As Result As
  Begin
    Self := Json_Parser.Parse_List(Str);
    Return;
  End;

  Constructor Function Json_List(Str Clob) Return Self As Result As
  Begin
    Self := Json_Parser.Parse_List(Str);
    Return;
  End;

  Constructor Function Json_List(Cast Json_Value) Return Self As Result As
    x Number;
  Begin
    x := Cast.Object_Or_Array.Getobject(Self);
    Return;
  End;

  Member Procedure Append(Self In Out Nocopy Json_List, Elem Json_Value, Position Pls_Integer Default Null) As
    Indx         Pls_Integer;
    Insert_Value Json_Value := Nvl(Elem, Json_Value);
  Begin
    If (Position Is Null Or Position > Self.Count) Then
      --end of list
      Indx := Self.Count + 1;
      Self.List_Data.Extend(1);
      Self.List_Data(Indx) := Insert_Value;
    Elsif (Position < 1) Then
      --new first
      Indx := Self.Count;
      Self.List_Data.Extend(1);
      For x In Reverse 1 .. Indx Loop
        Self.List_Data(x + 1) := Self.List_Data(x);
      End Loop;
      Self.List_Data(1) := Insert_Value;
    Else
      Indx := Self.Count;
      Self.List_Data.Extend(1);
      For x In Reverse Position .. Indx Loop
        Self.List_Data(x + 1) := Self.List_Data(x);
      End Loop;
      Self.List_Data(Position) := Insert_Value;
    End If;
  
  End;

  Member Procedure Append(Self In Out Nocopy Json_List, Elem Varchar2, Position Pls_Integer Default Null) As
  Begin
    Append(Json_Value(Elem), Position);
  End;

  Member Procedure Append(Self In Out Nocopy Json_List, Elem Number, Position Pls_Integer Default Null) As
  Begin
    If (Elem Is Null) Then
      Append(Json_Value(), Position);
    Else
      Append(Json_Value(Elem), Position);
    End If;
  End;

  Member Procedure Append(Self In Out Nocopy Json_List, Elem Boolean, Position Pls_Integer Default Null) As
  Begin
    If (Elem Is Null) Then
      Append(Json_Value(), Position);
    Else
      Append(Json_Value(Elem), Position);
    End If;
  End;

  Member Procedure Append(Self In Out Nocopy Json_List, Elem Json_List, Position Pls_Integer Default Null) As
  Begin
    If (Elem Is Null) Then
      Append(Json_Value(), Position);
    Else
      Append(Elem.To_Json_Value, Position);
    End If;
  End;

  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_Value) As
    Insert_Value Json_Value := Nvl(Elem, Json_Value);
    Indx         Number;
  Begin
    If (Position > Self.Count) Then
      --end of list
      Indx := Self.Count + 1;
      Self.List_Data.Extend(1);
      Self.List_Data(Indx) := Insert_Value;
    Elsif (Position < 1) Then
      --maybe an error message here
      Null;
    Else
      Self.List_Data(Position) := Insert_Value;
    End If;
  End;

  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Varchar2) As
  Begin
    Replace(Position, Json_Value(Elem));
  End;

  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Number) As
  Begin
    If (Elem Is Null) Then
      Replace(Position, Json_Value());
    Else
      Replace(Position, Json_Value(Elem));
    End If;
  End;

  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Boolean) As
  Begin
    If (Elem Is Null) Then
      Replace(Position, Json_Value());
    Else
      Replace(Position, Json_Value(Elem));
    End If;
  End;

  Member Procedure Replace(Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_List) As
  Begin
    If (Elem Is Null) Then
      Replace(Position, Json_Value());
    Else
      Replace(Position, Elem.To_Json_Value);
    End If;
  End;

  Member Function Count Return Number As
  Begin
    Return Self.List_Data.Count;
  End;

  Member Procedure Remove(Self In Out Nocopy Json_List, Position Pls_Integer) As
  Begin
    If (Position Is Null Or Position < 1 Or Position > Self.Count) Then
      Return;
    End If;
    For x In (Position + 1) .. Self.Count Loop
      Self.List_Data(x - 1) := Self.List_Data(x);
    End Loop;
    Self.List_Data.Trim(1);
  End;

  Member Procedure Remove_First(Self In Out Nocopy Json_List) As
  Begin
    For x In 2 .. Self.Count Loop
      Self.List_Data(x - 1) := Self.List_Data(x);
    End Loop;
    If (Self.Count > 0) Then
      Self.List_Data.Trim(1);
    End If;
  End;

  Member Procedure Remove_Last(Self In Out Nocopy Json_List) As
  Begin
    If (Self.Count > 0) Then
      Self.List_Data.Trim(1);
    End If;
  End;

  Member Function Get(Position Pls_Integer) Return Json_Value As
  Begin
    If (Self.Count >= Position And Position > 0) Then
      Return Self.List_Data(Position);
    End If;
    Return Null; -- do not throw error, just return null
  End;

  Member Function Head Return Json_Value As
  Begin
    If (Self.Count > 0) Then
      Return Self.List_Data(Self.List_Data.First);
    End If;
    Return Null; -- do not throw error, just return null
  End;

  Member Function Last Return Json_Value As
  Begin
    If (Self.Count > 0) Then
      Return Self.List_Data(Self.List_Data.Last);
    End If;
    Return Null; -- do not throw error, just return null
  End;

  Member Function Tail Return Json_List As
    t Json_List;
  Begin
    If (Self.Count > 0) Then
      t := Json_List(Self.List_Data);
      t.Remove(1);
      Return t;
    Else
      Return Json_List();
    End If;
  End;

  Member Function To_Char(Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2 As
  Begin
    If (Spaces Is Null) Then
      Return Json_Printer.Pretty_Print_List(Self, Line_Length => Chars_Per_Line);
    Else
      Return Json_Printer.Pretty_Print_List(Self, Spaces, Line_Length => Chars_Per_Line);
    End If;
  End;

  Member Procedure To_Clob(Self In Json_List, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True) As
  Begin
    If (Spaces Is Null) Then
      Json_Printer.Pretty_Print_List(Self, False, Buf, Line_Length => Chars_Per_Line, Erase_Clob => Erase_Clob);
    Else
      Json_Printer.Pretty_Print_List(Self, Spaces, Buf, Line_Length => Chars_Per_Line, Erase_Clob => Erase_Clob);
    End If;
  End;

  Member Procedure Print(Self In Json_List, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null) As
    --32512 is the real maximum in sqldeveloper
    My_Clob Clob;
  Begin
    My_Clob := Empty_Clob();
    Dbms_Lob.Createtemporary(My_Clob, True);
    Json_Printer.Pretty_Print_List(Self, Spaces, My_Clob, Case When(Chars_Per_Line > 32512) Then 32512 Else
                                    Chars_Per_Line End);
    Json_Printer.Dbms_Output_Clob(My_Clob, Json_Printer.Newline_Char, Jsonp);
    Dbms_Lob.Freetemporary(My_Clob);
  End;

  Member Procedure Htp(Self In Json_List, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null) As
    My_Clob Clob;
  Begin
    My_Clob := Empty_Clob();
    Dbms_Lob.Createtemporary(My_Clob, True);
    Json_Printer.Pretty_Print_List(Self, Spaces, My_Clob, Chars_Per_Line);
    Json_Printer.Htp_Output_Clob(My_Clob, Jsonp);
    Dbms_Lob.Freetemporary(My_Clob);
  End;

  /* json path */
  Member Function Path(Json_Path Varchar2, Base Number Default 1) Return Json_Value As
    Cp Json_List := Self;
  Begin
    Return Json_Ext.Get_Json_Value(Json(Cp), Json_Path, Base);
  End Path;

  /* json path_put */
  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_Value, Base Number Default 1) As
    Objlist Json;
    Jp      Json_List := Json_Ext.Parsepath(Json_Path, Base);
  Begin
    While (Jp.Head().Get_Number() > Self.Count) Loop
      Self.Append(Json_Value());
    End Loop;
  
    Objlist := Json(Self);
    Json_Ext.Put(Objlist, Json_Path, Elem, Base);
    Self := Objlist.Get_Values;
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Varchar2, Base Number Default 1) As
    Objlist Json;
    Jp      Json_List := Json_Ext.Parsepath(Json_Path, Base);
  Begin
    While (Jp.Head().Get_Number() > Self.Count) Loop
      Self.Append(Json_Value());
    End Loop;
  
    Objlist := Json(Self);
    Json_Ext.Put(Objlist, Json_Path, Elem, Base);
    Self := Objlist.Get_Values;
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Number, Base Number Default 1) As
    Objlist Json;
    Jp      Json_List := Json_Ext.Parsepath(Json_Path, Base);
  Begin
    While (Jp.Head().Get_Number() > Self.Count) Loop
      Self.Append(Json_Value());
    End Loop;
  
    Objlist := Json(Self);
  
    If (Elem Is Null) Then
      Json_Ext.Put(Objlist, Json_Path, Json_Value, Base);
    Else
      Json_Ext.Put(Objlist, Json_Path, Elem, Base);
    End If;
    Self := Objlist.Get_Values;
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Boolean, Base Number Default 1) As
    Objlist Json;
    Jp      Json_List := Json_Ext.Parsepath(Json_Path, Base);
  Begin
    While (Jp.Head().Get_Number() > Self.Count) Loop
      Self.Append(Json_Value());
    End Loop;
  
    Objlist := Json(Self);
    If (Elem Is Null) Then
      Json_Ext.Put(Objlist, Json_Path, Json_Value, Base);
    Else
      Json_Ext.Put(Objlist, Json_Path, Elem, Base);
    End If;
    Self := Objlist.Get_Values;
  End Path_Put;

  Member Procedure Path_Put(Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_List, Base Number Default 1) As
    Objlist Json;
    Jp      Json_List := Json_Ext.Parsepath(Json_Path, Base);
  Begin
    While (Jp.Head().Get_Number() > Self.Count) Loop
      Self.Append(Json_Value());
    End Loop;
  
    Objlist := Json(Self);
    If (Elem Is Null) Then
      Json_Ext.Put(Objlist, Json_Path, Json_Value, Base);
    Else
      Json_Ext.Put(Objlist, Json_Path, Elem, Base);
    End If;
    Self := Objlist.Get_Values;
  End Path_Put;

  /* json path_remove */
  Member Procedure Path_Remove(Self In Out Nocopy Json_List, Json_Path Varchar2, Base Number Default 1) As
    Objlist Json := Json(Self);
  Begin
    Json_Ext.Remove(Objlist, Json_Path, Base);
    Self := Objlist.Get_Values;
  End Path_Remove;

  Member Function To_Json_Value Return Json_Value As
  Begin
    Return Json_Value(Sys.Anydata.Convertobject(Self));
  End;
End;
/

Create Or Replace Package Json_Ac As
  --json type methods

  Procedure Object_Remove(p_Self In Out Nocopy Json, Pair_Name Varchar2);
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_Value, Position Pls_Integer Default Null);
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Varchar2, Position Pls_Integer Default Null);
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Number, Position Pls_Integer Default Null);
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Boolean, Position Pls_Integer Default Null);
  Procedure Object_Check_Duplicate(p_Self In Out Nocopy Json, v_Set Boolean);
  Procedure Object_Remove_Duplicates(p_Self In Out Nocopy Json);

  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json, Position Pls_Integer Default Null);
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_List, Position Pls_Integer Default Null);

  Function Object_Count(p_Self In Json) Return Number;
  Function Object_Get(p_Self In Json, Pair_Name Varchar2) Return Json_Value;
  Function Object_Get(p_Self In Json, Position Pls_Integer) Return Json_Value;
  Function Object_Index_Of(p_Self In Json, Pair_Name Varchar2) Return Number;
  Function Object_Exist(p_Self In Json, Pair_Name Varchar2) Return Boolean;

  Function Object_To_Char(p_Self In Json, Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2;
  Procedure Object_To_Clob(p_Self In Json, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True);
  Procedure Object_Print(p_Self In Json, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null);
  Procedure Object_Htp(p_Self In Json, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null);

  Function Object_To_Json_Value(p_Self In Json) Return Json_Value;
  Function Object_Path(p_Self In Json, Json_Path Varchar2, Base Number Default 1) Return Json_Value;

  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_Value, Base Number Default 1);
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Varchar2, Base Number Default 1);
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Number, Base Number Default 1);
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Boolean, Base Number Default 1);
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_List, Base Number Default 1);
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Json, Base Number Default 1);

  Procedure Object_Path_Remove(p_Self In Out Nocopy Json, Json_Path Varchar2, Base Number Default 1);

  Function Object_Get_Values(p_Self In Json) Return Json_List;
  Function Object_Get_Keys(p_Self In Json) Return Json_List;

  --json_list type methods  
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Json_Value, Position Pls_Integer Default Null);
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Varchar2, Position Pls_Integer Default Null);
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Number, Position Pls_Integer Default Null);
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Boolean, Position Pls_Integer Default Null);
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Json_List, Position Pls_Integer Default Null);

  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_Value);
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Varchar2);
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Number);
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Boolean);
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_List);

  Function Array_Count(p_Self In Json_List) Return Number;
  Procedure Array_Remove(p_Self In Out Nocopy Json_List, Position Pls_Integer);
  Procedure Array_Remove_First(p_Self In Out Nocopy Json_List);
  Procedure Array_Remove_Last(p_Self In Out Nocopy Json_List);
  Function Array_Get(p_Self In Json_List, Position Pls_Integer) Return Json_Value;
  Function Array_Head(p_Self In Json_List) Return Json_Value;
  Function Array_Last(p_Self In Json_List) Return Json_Value;
  Function Array_Tail(p_Self In Json_List) Return Json_List;

  Function Array_To_Char(p_Self In Json_List, Spaces Boolean Default True, Chars_Per_Line Number Default 0)
    Return Varchar2;
  Procedure Array_To_Clob(p_Self In Json_List, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True);
  Procedure Array_Print(p_Self In Json_List, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null);
  Procedure Array_Htp(p_Self In Json_List, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null);

  Function Array_Path(p_Self In Json_List, Json_Path Varchar2, Base Number Default 1) Return Json_Value;
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_Value, Base Number Default 1);
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Varchar2, Base Number Default 1);
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Number, Base Number Default 1);
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Boolean, Base Number Default 1);
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_List, Base Number Default 1);

  Procedure Array_Path_Remove(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Base Number Default 1);

  Function Array_To_Json_Value(p_Self In Json_List) Return Json_Value;

  --json_value

  Function Jv_Get_Type(p_Self In Json_Value) Return Varchar2;
  Function Jv_Get_String(p_Self In Json_Value, Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null)
    Return Varchar2;
  Procedure Jv_Get_String(p_Self In Json_Value, Buf In Out Nocopy Clob);
  Function Jv_Get_Number(p_Self In Json_Value) Return Number;
  Function Jv_Get_Bool(p_Self In Json_Value) Return Boolean;
  Function Jv_Get_Null(p_Self In Json_Value) Return Varchar2;

  Function Jv_Is_Object(p_Self In Json_Value) Return Boolean;
  Function Jv_Is_Array(p_Self In Json_Value) Return Boolean;
  Function Jv_Is_String(p_Self In Json_Value) Return Boolean;
  Function Jv_Is_Number(p_Self In Json_Value) Return Boolean;
  Function Jv_Is_Bool(p_Self In Json_Value) Return Boolean;
  Function Jv_Is_Null(p_Self In Json_Value) Return Boolean;

  Function Jv_To_Char(p_Self In Json_Value, Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2;
  Procedure Jv_To_Clob(p_Self In Json_Value, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True);
  Procedure Jv_Print(p_Self In Json_Value, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null);
  Procedure Jv_Htp(p_Self In Json_Value, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null);

  Function Jv_Value_Of(p_Self In Json_Value, Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null)
    Return Varchar2;

End Json_Ac;
/

Create Or Replace Package Body Json_Ac As
  Procedure Object_Remove(p_Self In Out Nocopy Json, Pair_Name Varchar2) As
  Begin
    p_Self.Remove(Pair_Name);
  End;
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_Value, Position Pls_Integer Default Null) As
  Begin
    p_Self.Put(Pair_Name, Pair_Value, Position);
  End;
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Varchar2, Position Pls_Integer Default Null) As
  Begin
    p_Self.Put(Pair_Name, Pair_Value, Position);
  End;
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Number, Position Pls_Integer Default Null) As
  Begin
    p_Self.Put(Pair_Name, Pair_Value, Position);
  End;
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Boolean, Position Pls_Integer Default Null) As
  Begin
    p_Self.Put(Pair_Name, Pair_Value, Position);
  End;
  Procedure Object_Check_Duplicate(p_Self In Out Nocopy Json, v_Set Boolean) As
  Begin
    p_Self.Check_Duplicate(v_Set);
  End;
  Procedure Object_Remove_Duplicates(p_Self In Out Nocopy Json) As
  Begin
    p_Self.Remove_Duplicates;
  End;

  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json, Position Pls_Integer Default Null) As
  Begin
    p_Self.Put(Pair_Name, Pair_Value, Position);
  End;
  Procedure Object_Put(p_Self In Out Nocopy Json, Pair_Name Varchar2, Pair_Value Json_List, Position Pls_Integer Default Null) As
  Begin
    p_Self.Put(Pair_Name, Pair_Value, Position);
  End;

  Function Object_Count(p_Self In Json) Return Number As
  Begin
    Return p_Self.Count;
  End;
  Function Object_Get(p_Self In Json, Pair_Name Varchar2) Return Json_Value As
  Begin
    Return p_Self.Get(Pair_Name);
  End;
  Function Object_Get(p_Self In Json, Position Pls_Integer) Return Json_Value As
  Begin
    Return p_Self.Get(Position);
  End;
  Function Object_Index_Of(p_Self In Json, Pair_Name Varchar2) Return Number As
  Begin
    Return p_Self.Index_Of(Pair_Name);
  End;
  Function Object_Exist(p_Self In Json, Pair_Name Varchar2) Return Boolean As
  Begin
    Return p_Self.Exist(Pair_Name);
  End;

  Function Object_To_Char(p_Self In Json, Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2 As
  Begin
    Return p_Self.To_Char(Spaces, Chars_Per_Line);
  End;
  Procedure Object_To_Clob(p_Self In Json, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True) As
  Begin
    p_Self.To_Clob(Buf, Spaces, Chars_Per_Line, Erase_Clob);
  End;
  Procedure Object_Print(p_Self In Json, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null) As
  Begin
    p_Self.Print(Spaces, Chars_Per_Line, Jsonp);
  End;
  Procedure Object_Htp(p_Self In Json, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null) As
  Begin
    p_Self.Htp(Spaces, Chars_Per_Line, Jsonp);
  End;

  Function Object_To_Json_Value(p_Self In Json) Return Json_Value As
  Begin
    Return p_Self.To_Json_Value;
  End;
  Function Object_Path(p_Self In Json, Json_Path Varchar2, Base Number Default 1) Return Json_Value As
  Begin
    Return p_Self.Path(Json_Path, Base);
  End;

  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_Value, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Varchar2, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Number, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Boolean, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Json_List, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Object_Path_Put(p_Self In Out Nocopy Json, Json_Path Varchar2, Elem Json, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;

  Procedure Object_Path_Remove(p_Self In Out Nocopy Json, Json_Path Varchar2, Base Number Default 1) As
  Begin
    p_Self.Path_Remove(Json_Path, Base);
  End;

  Function Object_Get_Values(p_Self In Json) Return Json_List As
  Begin
    Return p_Self.Get_Values;
  End;
  Function Object_Get_Keys(p_Self In Json) Return Json_List As
  Begin
    Return p_Self.Get_Keys;
  End;

  --json_list type
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Json_Value, Position Pls_Integer Default Null) As
  Begin
    p_Self.Append(Elem, Position);
  End;
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Varchar2, Position Pls_Integer Default Null) As
  Begin
    p_Self.Append(Elem, Position);
  End;
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Number, Position Pls_Integer Default Null) As
  Begin
    p_Self.Append(Elem, Position);
  End;
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Boolean, Position Pls_Integer Default Null) As
  Begin
    p_Self.Append(Elem, Position);
  End;
  Procedure Array_Append(p_Self In Out Nocopy Json_List, Elem Json_List, Position Pls_Integer Default Null) As
  Begin
    p_Self.Append(Elem, Position);
  End;

  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_Value) As
  Begin
    p_Self.Replace(Position, Elem);
  End;
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Varchar2) As
  Begin
    p_Self.Replace(Position, Elem);
  End;
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Number) As
  Begin
    p_Self.Replace(Position, Elem);
  End;
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Boolean) As
  Begin
    p_Self.Replace(Position, Elem);
  End;
  Procedure Array_Replace(p_Self In Out Nocopy Json_List, Position Pls_Integer, Elem Json_List) As
  Begin
    p_Self.Replace(Position, Elem);
  End;

  Function Array_Count(p_Self In Json_List) Return Number As
  Begin
    Return p_Self.Count;
  End;
  Procedure Array_Remove(p_Self In Out Nocopy Json_List, Position Pls_Integer) As
  Begin
    p_Self.Remove(Position);
  End;
  Procedure Array_Remove_First(p_Self In Out Nocopy Json_List) As
  Begin
    p_Self.Remove_First;
  End;
  Procedure Array_Remove_Last(p_Self In Out Nocopy Json_List) As
  Begin
    p_Self.Remove_Last;
  End;
  Function Array_Get(p_Self In Json_List, Position Pls_Integer) Return Json_Value As
  Begin
    Return p_Self.Get(Position);
  End;
  Function Array_Head(p_Self In Json_List) Return Json_Value As
  Begin
    Return p_Self.Head;
  End;
  Function Array_Last(p_Self In Json_List) Return Json_Value As
  Begin
    Return p_Self.Last;
  End;
  Function Array_Tail(p_Self In Json_List) Return Json_List As
  Begin
    Return p_Self.Tail;
  End;

  Function Array_To_Char(p_Self In Json_List, Spaces Boolean Default True, Chars_Per_Line Number Default 0)
    Return Varchar2 As
  Begin
    Return p_Self.To_Char(Spaces, Chars_Per_Line);
  End;
  Procedure Array_To_Clob(p_Self In Json_List, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True) As
  Begin
    p_Self.To_Clob(Buf, Spaces, Chars_Per_Line, Erase_Clob);
  End;
  Procedure Array_Print(p_Self In Json_List, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null) As
  Begin
    p_Self.Print(Spaces, Chars_Per_Line, Jsonp);
  End;
  Procedure Array_Htp(p_Self In Json_List, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null) As
  Begin
    p_Self.Htp(Spaces, Chars_Per_Line, Jsonp);
  End;

  Function Array_Path(p_Self In Json_List, Json_Path Varchar2, Base Number Default 1) Return Json_Value As
  Begin
    Return p_Self.Path(Json_Path, Base);
  End;
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_Value, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Varchar2, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Number, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Boolean, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;
  Procedure Array_Path_Put(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Elem Json_List, Base Number Default 1) As
  Begin
    p_Self.Path_Put(Json_Path, Elem, Base);
  End;

  Procedure Array_Path_Remove(p_Self In Out Nocopy Json_List, Json_Path Varchar2, Base Number Default 1) As
  Begin
    p_Self.Path_Remove(Json_Path, Base);
  End;

  Function Array_To_Json_Value(p_Self In Json_List) Return Json_Value As
  Begin
    Return p_Self.To_Json_Value;
  End;

  --json_value

  Function Jv_Get_Type(p_Self In Json_Value) Return Varchar2 As
  Begin
    Return p_Self.Get_Type;
  End;
  Function Jv_Get_String(p_Self In Json_Value, Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null)
    Return Varchar2 As
  Begin
    Return p_Self.Get_String(Max_Byte_Size, Max_Char_Size);
  End;
  Procedure Jv_Get_String(p_Self In Json_Value, Buf In Out Nocopy Clob) As
  Begin
    p_Self.Get_String(Buf);
  End;
  Function Jv_Get_Number(p_Self In Json_Value) Return Number As
  Begin
    Return p_Self.Get_Number;
  End;
  Function Jv_Get_Bool(p_Self In Json_Value) Return Boolean As
  Begin
    Return p_Self.Get_Bool;
  End;
  Function Jv_Get_Null(p_Self In Json_Value) Return Varchar2 As
  Begin
    Return p_Self.Get_Null;
  End;

  Function Jv_Is_Object(p_Self In Json_Value) Return Boolean As
  Begin
    Return p_Self.Is_Object;
  End;
  Function Jv_Is_Array(p_Self In Json_Value) Return Boolean As
  Begin
    Return p_Self.Is_Array;
  End;
  Function Jv_Is_String(p_Self In Json_Value) Return Boolean As
  Begin
    Return p_Self.Is_String;
  End;
  Function Jv_Is_Number(p_Self In Json_Value) Return Boolean As
  Begin
    Return p_Self.Is_Number;
  End;
  Function Jv_Is_Bool(p_Self In Json_Value) Return Boolean As
  Begin
    Return p_Self.Is_Bool;
  End;
  Function Jv_Is_Null(p_Self In Json_Value) Return Boolean As
  Begin
    Return p_Self.Is_Null;
  End;

  Function Jv_To_Char(p_Self In Json_Value, Spaces Boolean Default True, Chars_Per_Line Number Default 0) Return Varchar2 As
  Begin
    Return p_Self.To_Char(Spaces, Chars_Per_Line);
  End;
  Procedure Jv_To_Clob(p_Self In Json_Value, Buf In Out Nocopy Clob, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Erase_Clob Boolean Default True) As
  Begin
    p_Self.To_Clob(Buf, Spaces, Chars_Per_Line, Erase_Clob);
  End;
  Procedure Jv_Print(p_Self In Json_Value, Spaces Boolean Default True, Chars_Per_Line Number Default 8192, Jsonp Varchar2 Default Null) As
  Begin
    p_Self.Print(Spaces, Chars_Per_Line, Jsonp);
  End;
  Procedure Jv_Htp(p_Self In Json_Value, Spaces Boolean Default False, Chars_Per_Line Number Default 0, Jsonp Varchar2 Default Null) As
  Begin
    p_Self.Htp(Spaces, Chars_Per_Line, Jsonp);
  End;

  Function Jv_Value_Of(p_Self In Json_Value, Max_Byte_Size Number Default Null, Max_Char_Size Number Default Null)
    Return Varchar2 As
  Begin
    Return p_Self.Value_Of(Max_Byte_Size, Max_Char_Size);
  End;

End Json_Ac;
/

Create Or Replace Package Json_Dyn Authid Current_User As
  Null_As_Empty_String Boolean Not Null := True; --varchar2
  Include_Dates        Boolean Not Null := True; --date
  Include_Clobs        Boolean Not Null := True;
  Include_Blobs        Boolean Not Null := False;

  /* list with objects */
  Function Executelist(Stmt Varchar2, Bindvar Json Default Null, Cur_Num Number Default Null) Return Json_List;

  /* object with lists */
  Function Executeobject(Stmt Varchar2, Bindvar Json Default Null, Cur_Num Number Default Null) Return Json;
End Json_Dyn;
/

Create Or Replace Package Body Json_Dyn As
  Procedure Bind_Json(l_Cur Number, Bindvar Json) As
    Keylist Json_List := Bindvar.Get_Keys();
  Begin
    For i In 1 .. Keylist.Count Loop
      If (Bindvar.Get(i).Get_Type = 'number') Then
        Dbms_Sql.Bind_Variable(l_Cur, ':' || Keylist.Get(i).Get_String, Bindvar.Get(i).Get_Number);
      Elsif (Bindvar.Get(i).Get_Type = 'array') Then
        Declare
          v_Bind Dbms_Sql.Varchar2_Table;
          v_Arr  Json_List := Json_List(Bindvar.Get(i));
        Begin
          For j In 1 .. v_Arr.Count Loop
            v_Bind(j) := v_Arr.Get(j).Value_Of;
          End Loop;
          Dbms_Sql.Bind_Array(l_Cur, ':' || Keylist.Get(i).Get_String, v_Bind);
        End;
      Else
        Dbms_Sql.Bind_Variable(l_Cur, ':' || Keylist.Get(i).Get_String, Bindvar.Get(i).Value_Of());
      End If;
    End Loop;
  End Bind_Json;

  /* list with objects */
  Function Executelist(Stmt Varchar2, Bindvar Json, Cur_Num Number) Return Json_List As
    l_Cur      Number;
    l_Dtbl     Dbms_Sql.Desc_Tab2;
    l_Cnt      Number;
    l_Status   Number;
    l_Val      Varchar2(4000);
    Outer_List Json_List := Json_List();
    Inner_Obj  Json;
    Conv       Number;
    Read_Date  Date;
    Read_Clob  Clob;
    Read_Blob  Blob;
    Col_Type   Number;
  Begin
    If (Cur_Num Is Not Null) Then
      l_Cur := Cur_Num;
    Else
      l_Cur := Dbms_Sql.Open_Cursor;
      Dbms_Sql.Parse(l_Cur, Stmt, Dbms_Sql.Native);
      If (Bindvar Is Not Null) Then
        Bind_Json(l_Cur, Bindvar);
      End If;
    End If;
    Dbms_Sql.Describe_Columns2(l_Cur, l_Cnt, l_Dtbl);
    For i In 1 .. l_Cnt Loop
      Col_Type := l_Dtbl(i).Col_Type;
      --dbms_output.put_line(col_type);
      If (Col_Type = 12) Then
        Dbms_Sql.Define_Column(l_Cur, i, Read_Date);
      Elsif (Col_Type = 112) Then
        Dbms_Sql.Define_Column(l_Cur, i, Read_Clob);
      Elsif (Col_Type = 113) Then
        Dbms_Sql.Define_Column(l_Cur, i, Read_Blob);
      Elsif (Col_Type In (1, 2, 96)) Then
        Dbms_Sql.Define_Column(l_Cur, i, l_Val, 4000);
      End If;
    End Loop;
  
    If (Cur_Num Is Null) Then
      l_Status := Dbms_Sql.Execute(l_Cur);
    End If;
  
    --loop through rows
    While (Dbms_Sql.Fetch_Rows(l_Cur) > 0) Loop
      Inner_Obj := Json(); --init for each row
      --loop through columns
      For i In 1 .. l_Cnt Loop
        Case True
        --handling string types
          When l_Dtbl(i).Col_Type In (1, 96) Then
            -- varchar2
            Dbms_Sql.Column_Value(l_Cur, i, l_Val);
            If (l_Val Is Null) Then
              If (Null_As_Empty_String) Then
                Inner_Obj.Put(l_Dtbl(i).Col_Name, ''); --treatet as emptystring
              Else
                Inner_Obj.Put(l_Dtbl(i).Col_Name, Json_Value.Makenull); --null
              End If;
            Else
              Inner_Obj.Put(l_Dtbl(i).Col_Name, Json_Value(l_Val)); --null
            End If;
            --handling number types
          When l_Dtbl(i).Col_Type = 2 Then
            -- number
            Dbms_Sql.Column_Value(l_Cur, i, l_Val);
            Conv := l_Val;
            Inner_Obj.Put(l_Dtbl(i).Col_Name, Conv);
          When l_Dtbl(i).Col_Type = 12 Then
            -- date
            If (Include_Dates) Then
              Dbms_Sql.Column_Value(l_Cur, i, Read_Date);
              Inner_Obj.Put(l_Dtbl(i).Col_Name, Json_Ext.To_Json_Value(Read_Date));
            End If;
          When l_Dtbl(i).Col_Type = 112 Then
            --clob
            If (Include_Clobs) Then
              Dbms_Sql.Column_Value(l_Cur, i, Read_Clob);
              Inner_Obj.Put(l_Dtbl(i).Col_Name, Json_Value(Read_Clob));
            End If;
          When l_Dtbl(i).Col_Type = 113 Then
            --blob
            If (Include_Blobs) Then
              Dbms_Sql.Column_Value(l_Cur, i, Read_Blob);
              If (Dbms_Lob.Getlength(Read_Blob) > 0) Then
                Inner_Obj.Put(l_Dtbl(i).Col_Name, Json_Ext.Encode(Read_Blob));
              Else
                Inner_Obj.Put(l_Dtbl(i).Col_Name, Json_Value.Makenull);
              End If;
            End If;
          
          Else
            Null; --discard other types
        End Case;
      End Loop;
      Outer_List.Append(Inner_Obj.To_Json_Value);
    End Loop;
    Dbms_Sql.Close_Cursor(l_Cur);
    Return Outer_List;
  End Executelist;

  /* object with lists */
  Function Executeobject(Stmt Varchar2, Bindvar Json, Cur_Num Number) Return Json As
    l_Cur            Number;
    l_Dtbl           Dbms_Sql.Desc_Tab;
    l_Cnt            Number;
    l_Status         Number;
    l_Val            Varchar2(4000);
    Inner_List_Names Json_List := Json_List();
    Inner_List_Data  Json_List := Json_List();
    Data_List        Json_List;
    Outer_Obj        Json := Json();
    Conv             Number;
    Read_Date        Date;
    Read_Clob        Clob;
    Read_Blob        Blob;
    Col_Type         Number;
  Begin
    If (Cur_Num Is Not Null) Then
      l_Cur := Cur_Num;
    Else
      l_Cur := Dbms_Sql.Open_Cursor;
      Dbms_Sql.Parse(l_Cur, Stmt, Dbms_Sql.Native);
      If (Bindvar Is Not Null) Then
        Bind_Json(l_Cur, Bindvar);
      End If;
    End If;
    Dbms_Sql.Describe_Columns(l_Cur, l_Cnt, l_Dtbl);
    For i In 1 .. l_Cnt Loop
      Col_Type := l_Dtbl(i).Col_Type;
      If (Col_Type = 12) Then
        Dbms_Sql.Define_Column(l_Cur, i, Read_Date);
      Elsif (Col_Type = 112) Then
        Dbms_Sql.Define_Column(l_Cur, i, Read_Clob);
      Elsif (Col_Type = 113) Then
        Dbms_Sql.Define_Column(l_Cur, i, Read_Blob);
      Elsif (Col_Type In (1, 2, 96)) Then
        Dbms_Sql.Define_Column(l_Cur, i, l_Val, 4000);
      End If;
    End Loop;
    If (Cur_Num Is Null) Then
      l_Status := Dbms_Sql.Execute(l_Cur);
    End If;
  
    --build up name_list
    For i In 1 .. l_Cnt Loop
      Case l_Dtbl(i).Col_Type
        When 1 Then
          Inner_List_Names.Append(l_Dtbl(i).Col_Name);
        When 96 Then
          Inner_List_Names.Append(l_Dtbl(i).Col_Name);
        When 2 Then
          Inner_List_Names.Append(l_Dtbl(i).Col_Name);
        When 12 Then
          If (Include_Dates) Then
            Inner_List_Names.Append(l_Dtbl(i).Col_Name);
          End If;
        When 112 Then
          If (Include_Clobs) Then
            Inner_List_Names.Append(l_Dtbl(i).Col_Name);
          End If;
        When 113 Then
          If (Include_Blobs) Then
            Inner_List_Names.Append(l_Dtbl(i).Col_Name);
          End If;
        Else
          Null;
      End Case;
    End Loop;
  
    --loop through rows
    While (Dbms_Sql.Fetch_Rows(l_Cur) > 0) Loop
      Data_List := Json_List();
      --loop through columns
      For i In 1 .. l_Cnt Loop
        Case True
        --handling string types
          When l_Dtbl(i).Col_Type In (1, 96) Then
            -- varchar2
            Dbms_Sql.Column_Value(l_Cur, i, l_Val);
            If (l_Val Is Null) Then
              If (Null_As_Empty_String) Then
                Data_List.Append(''); --treatet as emptystring
              Else
                Data_List.Append(Json_Value.Makenull); --null
              End If;
            Else
              Data_List.Append(Json_Value(l_Val)); --null
            End If;
            --handling number types
          When l_Dtbl(i).Col_Type = 2 Then
            -- number
            Dbms_Sql.Column_Value(l_Cur, i, l_Val);
            Conv := l_Val;
            Data_List.Append(Conv);
          When l_Dtbl(i).Col_Type = 12 Then
            -- date
            If (Include_Dates) Then
              Dbms_Sql.Column_Value(l_Cur, i, Read_Date);
              Data_List.Append(Json_Ext.To_Json_Value(Read_Date));
            End If;
          When l_Dtbl(i).Col_Type = 112 Then
            --clob
            If (Include_Clobs) Then
              Dbms_Sql.Column_Value(l_Cur, i, Read_Clob);
              Data_List.Append(Json_Value(Read_Clob));
            End If;
          When l_Dtbl(i).Col_Type = 113 Then
            --blob
            If (Include_Blobs) Then
              Dbms_Sql.Column_Value(l_Cur, i, Read_Blob);
              If (Dbms_Lob.Getlength(Read_Blob) > 0) Then
                Data_List.Append(Json_Ext.Encode(Read_Blob));
              Else
                Data_List.Append(Json_Value.Makenull);
              End If;
            End If;
          Else
            Null; --discard other types
        End Case;
      End Loop;
      Inner_List_Data.Append(Data_List);
    End Loop;
  
    Outer_Obj.Put('names', Inner_List_Names.To_Json_Value);
    Outer_Obj.Put('data', Inner_List_Data.To_Json_Value);
    Dbms_Sql.Close_Cursor(l_Cur);
    Return Outer_Obj;
  End Executeobject;

End Json_Dyn;
/

Create Or Replace Package Json_Ml As
  Jsonml_Stylesheet Xmltype := Null;

  Function Xml2json(Xml In Xmltype) Return Json_List;
  Function Xmlstr2json(Xmlstr In Varchar2) Return Json_List;

End Json_Ml;
/


Create Or Replace Package Body Json_Ml As
  Function Get_Jsonml_Stylesheet Return Xmltype;

  Function Xml2json(Xml In Xmltype) Return Json_List As
    l_Json        Xmltype;
    l_Returnvalue Clob;
  Begin
    l_Json        := Xml.Transform(Get_Jsonml_Stylesheet);
    l_Returnvalue := l_Json.Getclobval();
    l_Returnvalue := Dbms_Xmlgen.Convert(l_Returnvalue, Dbms_Xmlgen.Entity_Decode);
    --dbms_output.put_line(l_returnvalue);
    Return Json_List(l_Returnvalue);
  End Xml2json;

  Function Xmlstr2json(Xmlstr In Varchar2) Return Json_List As
  Begin
    Return Xml2json(Xmltype(Xmlstr));
  End Xmlstr2json;

  Function Get_Jsonml_Stylesheet Return Xmltype As
  Begin
    If (Jsonml_Stylesheet Is Null) Then
      Jsonml_Stylesheet := Xmltype('<?xml version="1.0" encoding="UTF-8"?>
<!--
    JsonML.xslt
 
    Created: 2006-11-15-0551
    Modified: 2009-02-14-0927
 
    Released under an open-source license:
    http://jsonml.org/License.htm
 
    This transformation converts any XML document into JsonML.
    It omits processing-instructions and comment-nodes.
    
    To enable comment-nodes to be emitted as JavaScript comments,
    uncomment the Comment() template.
--> 
<xsl:stylesheet version="1.0"
        xmlns:xsl="http://www.w3.org/1999/XSL/Transform"> 
 
  <xsl:output method="text"
        media-type="application/json"
        encoding="UTF-8"
        indent="no"
        omit-xml-declaration="yes" /> 
 
  <!-- constants --> 
  <xsl:variable name="XHTML"
          select="''http://www.w3.org/1999/xhtml''" /> 
 
  <xsl:variable name="START_ELEM"
          select="''[''" /> 
 
  <xsl:variable name="END_ELEM"
          select="'']''" /> 
 
  <xsl:variable name="VALUE_DELIM"
          select="'',''" /> 
 
  <xsl:variable name="START_ATTRIB"
          select="''{''" /> 
 
  <xsl:variable name="END_ATTRIB"
          select="''}''" /> 
 
  <xsl:variable name="NAME_DELIM"
          select="'':''" /> 
 
  <xsl:variable name="STRING_DELIM"
          select="''' ||chr(38) ||'#x22;''" /> 
 
  <xsl:variable name="START_COMMENT"
          select="''/*''" /> 
 
  <xsl:variable name="END_COMMENT"
          select="''*/''" /> 
 
  <!-- root-node --> 
  <xsl:template match="/"> 
    <xsl:apply-templates select="*" /> 
  </xsl:template> 
 
  <!-- comments --> 
  <xsl:template match="comment()"> 
  <!-- uncomment to support JSON comments --> 
  <!--
    <xsl:value-of select="$START_COMMENT" />
 
    <xsl:value-of select="."
            disable-output-escaping="yes" />
 
    <xsl:value-of select="$END_COMMENT" />
  --> 
  </xsl:template> 
 
  <!-- elements --> 
  <xsl:template match="*"> 
    <xsl:value-of select="$START_ELEM" /> 
 
    <!-- tag-name string --> 
    <xsl:value-of select="$STRING_DELIM" /> 
    <xsl:choose> 
      <xsl:when test="namespace-uri()=$XHTML"> 
        <xsl:value-of select="local-name()" /> 
      </xsl:when> 
      <xsl:otherwise> 
        <xsl:value-of select="name()" /> 
      </xsl:otherwise> 
    </xsl:choose> 
    <xsl:value-of select="$STRING_DELIM" /> 
 
    <!-- attribute object --> 
    <xsl:if test="count(@*)>0"> 
      <xsl:value-of select="$VALUE_DELIM" /> 
      <xsl:value-of select="$START_ATTRIB" /> 
      <xsl:for-each select="@*"> 
        <xsl:if test="position()>1"> 
          <xsl:value-of select="$VALUE_DELIM" /> 
        </xsl:if> 
        <xsl:apply-templates select="." /> 
      </xsl:for-each> 
      <xsl:value-of select="$END_ATTRIB" /> 
    </xsl:if> 
 
    <!-- child elements and text-nodes --> 
    <xsl:for-each select="*|text()"> 
      <xsl:value-of select="$VALUE_DELIM" /> 
      <xsl:apply-templates select="." /> 
    </xsl:for-each> 
 
    <xsl:value-of select="$END_ELEM" /> 
  </xsl:template> 
 
  <!-- text-nodes --> 
  <xsl:template match="text()"> 
    <xsl:call-template name="escape-string"> 
      <xsl:with-param name="value"
              select="." /> 
    </xsl:call-template> 
  </xsl:template> 
 
  <!-- attributes --> 
  <xsl:template match="@*"> 
    <xsl:value-of select="$STRING_DELIM" /> 
    <xsl:choose> 
      <xsl:when test="namespace-uri()=$XHTML"> 
        <xsl:value-of select="local-name()" /> 
      </xsl:when> 
      <xsl:otherwise> 
        <xsl:value-of select="name()" /> 
      </xsl:otherwise> 
    </xsl:choose> 
    <xsl:value-of select="$STRING_DELIM" /> 
 
    <xsl:value-of select="$NAME_DELIM" /> 
 
    <xsl:call-template name="escape-string"> 
      <xsl:with-param name="value"
              select="." /> 
    </xsl:call-template> 
 
  </xsl:template> 
 
  <!-- escape-string: quotes and escapes --> 
  <xsl:template name="escape-string"> 
    <xsl:param name="value" /> 
 
    <xsl:value-of select="$STRING_DELIM" /> 
 
    <xsl:if test="string-length($value)>0"> 
      <xsl:variable name="escaped-whacks"> 
        <!-- escape backslashes --> 
        <xsl:call-template name="string-replace"> 
          <xsl:with-param name="value"
                  select="$value" /> 
          <xsl:with-param name="find"
                  select="''\''" /> 
          <xsl:with-param name="replace"
                  select="''\\''" /> 
        </xsl:call-template> 
      </xsl:variable> 
 
      <xsl:variable name="escaped-LF"> 
        <!-- escape line feeds --> 
        <xsl:call-template name="string-replace"> 
          <xsl:with-param name="value"
                  select="$escaped-whacks" /> 
          <xsl:with-param name="find"
                  select="''' ||chr(38) ||'#x0A;''" /> 
          <xsl:with-param name="replace"
                  select="''\n''" /> 
        </xsl:call-template> 
      </xsl:variable> 
 
      <xsl:variable name="escaped-CR"> 
        <!-- escape carriage returns --> 
        <xsl:call-template name="string-replace"> 
          <xsl:with-param name="value"
                  select="$escaped-LF" /> 
          <xsl:with-param name="find"
                  select="''' ||chr(38) ||'#x0D;''" /> 
          <xsl:with-param name="replace"
                  select="''\r''" /> 
        </xsl:call-template> 
      </xsl:variable> 
 
      <xsl:variable name="escaped-tabs"> 
        <!-- escape tabs --> 
        <xsl:call-template name="string-replace"> 
          <xsl:with-param name="value"
                  select="$escaped-CR" /> 
          <xsl:with-param name="find"
                  select="''' ||chr(38) ||'#x09;''" /> 
          <xsl:with-param name="replace"
                  select="''\t''" /> 
        </xsl:call-template> 
      </xsl:variable> 
 
      <!-- escape quotes --> 
      <xsl:call-template name="string-replace"> 
        <xsl:with-param name="value"
                select="$escaped-tabs" /> 
        <xsl:with-param name="find"
                select="''' ||chr(38) ||'quot;''" /> 
        <xsl:with-param name="replace"
                select="''\' ||chr(38) ||'quot;''" /> 
      </xsl:call-template> 
    </xsl:if> 
 
    <xsl:value-of select="$STRING_DELIM" /> 
  </xsl:template> 
 
  <!-- string-replace: replaces occurances of one string with another --> 
  <xsl:template name="string-replace"> 
    <xsl:param name="value" /> 
    <xsl:param name="find" /> 
    <xsl:param name="replace" /> 
 
    <xsl:choose> 
      <xsl:when test="contains($value,$find)"> 
        <!-- replace and call recursively on next --> 
        <xsl:value-of select="substring-before($value,$find)"
                disable-output-escaping="yes" /> 
        <xsl:value-of select="$replace"
                disable-output-escaping="yes" /> 
        <xsl:call-template name="string-replace"> 
          <xsl:with-param name="value"
                  select="substring-after($value,$find)" /> 
          <xsl:with-param name="find"
                  select="$find" /> 
          <xsl:with-param name="replace"
                  select="$replace" /> 
        </xsl:call-template> 
      </xsl:when> 
      <xsl:otherwise> 
        <!-- no replacement necessary --> 
        <xsl:value-of select="$value"
                disable-output-escaping="yes" /> 
      </xsl:otherwise> 
    </xsl:choose> 
  </xsl:template> 
 
</xsl:stylesheet>');
    End If;
    Return Jsonml_Stylesheet;
  End Get_Jsonml_Stylesheet;

End Json_Ml;
/

Create Or Replace Package Json_Xml As

  Function Json_To_Xml(Obj Json, Tagname Varchar2 Default 'root') Return Xmltype;

End Json_Xml;
/

Create Or Replace Package Body Json_Xml As
  Function Escapestr(Str Varchar2) Return Varchar2 As
    Buf Varchar2(32767) := '';
    Ch  Varchar2(4);
  Begin
    For i In 1 .. Length(Str) Loop
      Ch := Substr(Str, i, 1);
      Case Ch
        When '' ||chr(38) ||'' Then
          Buf := Buf || '' ||chr(38) ||'amp;';
        When '<' Then
          Buf := Buf || '' ||chr(38) ||'lt;';
        When '>' Then
          Buf := Buf || '' ||chr(38) ||'gt;';
        When '"' Then
          Buf := Buf || '' ||chr(38) ||'quot;';
        Else
          Buf := Buf || Ch;
      End Case;
    End Loop;
    Return Buf;
  End Escapestr;

  /* Clob methods from printer */
  Procedure Add_To_Clob(Buf_Lob In Out Nocopy Clob, Buf_Str In Out Nocopy Varchar2, Str Varchar2) As
  Begin
    If (Length(Str) > 32767 - Length(Buf_Str)) Then
      Dbms_Lob.Append(Buf_Lob, Buf_Str);
      Buf_Str := Str;
    Else
      Buf_Str := Buf_Str || Str;
    End If;
  End Add_To_Clob;

  Procedure Flush_Clob(Buf_Lob In Out Nocopy Clob, Buf_Str In Out Nocopy Varchar2) As
  Begin
    Dbms_Lob.Append(Buf_Lob, Buf_Str);
  End Flush_Clob;

  Procedure Tostring(Obj Json_Value, Tagname In Varchar2, Xmlstr In Out Nocopy Clob, Xmlbuf In Out Nocopy Varchar2) As
    v_Obj  Json;
    v_List Json_List;
  
    v_Keys  Json_List;
    v_Value Json_Value;
    Key_Str Varchar2(4000);
  Begin
    If (Obj.Is_Object()) Then
      Add_To_Clob(Xmlstr, Xmlbuf, '<' || Tagname || '>');
      v_Obj := Json(Obj);
    
      v_Keys := v_Obj.Get_Keys();
      For i In 1 .. v_Keys.Count Loop
        v_Value := v_Obj.Get(i);
        Key_Str := v_Keys.Get(i).Str;
      
        If (Key_Str = 'content') Then
          If (v_Value.Is_Array()) Then
            Declare
              v_l Json_List := Json_List(v_Value);
            Begin
              For j In 1 .. v_l.Count Loop
                If (j > 1) Then
                  Add_To_Clob(Xmlstr, Xmlbuf, Chr(13) || Chr(10));
                End If;
                Add_To_Clob(Xmlstr, Xmlbuf, Escapestr(v_l.Get(j).To_Char()));
              End Loop;
            End;
          Else
            Add_To_Clob(Xmlstr, Xmlbuf, Escapestr(v_Value.To_Char()));
          End If;
        Elsif (v_Value.Is_Array()) Then
          Declare
            v_l Json_List := Json_List(v_Value);
          Begin
            For j In 1 .. v_l.Count Loop
              v_Value := v_l.Get(j);
              If (v_Value.Is_Array()) Then
                Add_To_Clob(Xmlstr, Xmlbuf, '<' || Key_Str || '>');
                Add_To_Clob(Xmlstr, Xmlbuf, Escapestr(v_Value.To_Char()));
                Add_To_Clob(Xmlstr, Xmlbuf, '</' || Key_Str || '>');
              Else
                Tostring(v_Value, Key_Str, Xmlstr, Xmlbuf);
              End If;
            End Loop;
          End;
        Elsif (v_Value.Is_Null() Or (v_Value.Is_String And v_Value.Get_String = '')) Then
          Add_To_Clob(Xmlstr, Xmlbuf, '<' || Key_Str || '/>');
        Else
          Tostring(v_Value, Key_Str, Xmlstr, Xmlbuf);
        End If;
      End Loop;
    
      Add_To_Clob(Xmlstr, Xmlbuf, '</' || Tagname || '>');
    Elsif (Obj.Is_Array()) Then
      v_List := Json_List(Obj);
      For i In 1 .. v_List.Count Loop
        v_Value := v_List.Get(i);
        Tostring(v_Value, Nvl(Tagname, 'array'), Xmlstr, Xmlbuf);
      End Loop;
    Else
      Add_To_Clob(Xmlstr, Xmlbuf, '<' || Tagname || '>' || Escapestr(Obj.Value_Of()) || '</' || Tagname || '>');
    End If;
  End Tostring;

  Function Json_To_Xml(Obj Json, Tagname Varchar2 Default 'root') Return Xmltype As
    Xmlstr      Clob := Empty_Clob();
    Xmlbuf      Varchar2(32767) := '';
    Returnvalue Xmltype;
  Begin
    Dbms_Lob.Createtemporary(Xmlstr, True);
  
    Tostring(Obj.To_Json_Value(), Tagname, Xmlstr, Xmlbuf);
  
    Flush_Clob(Xmlstr, Xmlbuf);
    Returnvalue := Xmltype('<?xml version="1.0"?>' || Xmlstr);
    Dbms_Lob.Freetemporary(Xmlstr);
    Return Returnvalue;
  End;

End Json_Xml;
/

Create Or Replace Package Json_Util_Pkg Authid Current_User As
  Function Ref_Cursor_To_Json(p_Ref_Cursor In Sys_Refcursor, p_Max_Rows In Number := Null, p_Skip_Rows In Number := Null)
    Return Json_List;

  -- generate JSON from SQL statement
  Function Sql_To_Json(p_Sql In Varchar2, p_Max_Rows In Number := Null, p_Skip_Rows In Number := Null) Return Json_List;

End Json_Util_Pkg;
/

Create Or Replace Package Body Json_Util_Pkg As
  Scanner_Exception Exception;
  Pragma Exception_Init(Scanner_Exception, -20100);
  Parser_Exception Exception;
  Pragma Exception_Init(Parser_Exception, -20101);
  g_Json_Null_Object Constant Varchar2(20) := '{ }';

  Function Get_Xml_To_Json_Stylesheet Return Varchar2 As
  Begin
    Return q'^<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<!--
  Copyright (c) 2006,2008 Doeke Zanstra
  All rights reserved.

  Redistribution and use in source and binary forms, with or without modification, 
  are permitted provided that the following conditions are met:

  Redistributions of source code must retain the above copyright notice, this 
  list of conditions and the following disclaimer. Redistributions in binary 
  form must reproduce the above copyright notice, this list of conditions and the 
  following disclaimer in the documentation and/or other materials provided with 
  the distribution.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND 
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
  IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
  INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
  DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
  LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR 
  OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF 
  THE POSSIBILITY OF SUCH DAMAGE.
-->

  <xsl:output indent="no" omit-xml-declaration="yes" method="text" encoding="UTF-8" media-type="text/x-json"/>
        <xsl:strip-space elements="*"/>
  <!--contant-->
  <xsl:variable name="d">0123456789</xsl:variable>

  <!-- ignore document text -->
  <xsl:template match="text()[preceding-sibling::node() or following-sibling::node()]"/>

  <!-- string -->
  <xsl:template match="text()">
    <xsl:call-template name="escape-string">
      <xsl:with-param name="s" select="."/>
    </xsl:call-template>
  </xsl:template>
  
  <!-- Main template for escaping strings; used by above template and for object-properties 
       Responsibilities: placed quotes around string, and chain up to next filter, escape-bs-string -->
  <xsl:template name="escape-string">
    <xsl:param name="s"/>
    <xsl:text>"</xsl:text>
    <xsl:call-template name="escape-bs-string">
      <xsl:with-param name="s" select="$s"/>
    </xsl:call-template>
    <xsl:text>"</xsl:text>
  </xsl:template>
  
  <!-- Escape the backslash (\) before everything else. -->
  <xsl:template name="escape-bs-string">
    <xsl:param name="s"/>
    <xsl:choose>
      <xsl:when test="contains($s,'\')">
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'\'),'\\')"/>
        </xsl:call-template>
        <xsl:call-template name="escape-bs-string">
          <xsl:with-param name="s" select="substring-after($s,'\')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="$s"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- Escape the double quote ("). -->
  <xsl:template name="escape-quot-string">
    <xsl:param name="s"/>
    <xsl:choose>
      <xsl:when test="contains($s,'' ||chr(38) ||'quot;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'' ||chr(38) ||'quot;'),'\' ||chr(38) ||'quot;')"/>
        </xsl:call-template>
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="substring-after($s,'' ||chr(38) ||'quot;')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="$s"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- Replace tab, line feed and/or carriage return by its matching escape code. Can''t escape backslash
       or double quote here, because they don''t replace characters (' ||chr(38) ||'#x0; becomes \t), but they prefix 
       characters (\ becomes \\). Besides, backslash should be seperate anyway, because it should be 
       processed first. This function can''t do that. -->
  <xsl:template name="encode-string">
    <xsl:param name="s"/>
    <xsl:choose>
      <!-- tab -->
      <xsl:when test="contains($s,'' ||chr(38) ||'#x9;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'' ||chr(38) ||'#x9;'),'\t',substring-after($s,'' ||chr(38) ||'#x9;'))"/>
        </xsl:call-template>
      </xsl:when>
      <!-- line feed -->
      <xsl:when test="contains($s,'' ||chr(38) ||'#xA;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'' ||chr(38) ||'#xA;'),'\n',substring-after($s,'' ||chr(38) ||'#xA;'))"/>
        </xsl:call-template>
      </xsl:when>
      <!-- carriage return -->
      <xsl:when test="contains($s,'' ||chr(38) ||'#xD;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'' ||chr(38) ||'#xD;'),'\r',substring-after($s,'' ||chr(38) ||'#xD;'))"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise><xsl:value-of select="$s"/></xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- number (no support for javascript mantissa) -->
  <xsl:template match="text()[not(string(number())='NaN' or
                      (starts-with(.,'0' ) and . != '0' and
not(starts-with(.,'0.' ))) or
                      (starts-with(.,'-0' ) and . != '-0' and
not(starts-with(.,'-0.' )))
                      )]">
    <xsl:value-of select="."/>
  </xsl:template>

  <!-- boolean, case-insensitive -->
  <xsl:template match="text()[translate(.,'TRUE','true')='true']">true</xsl:template>
  <xsl:template match="text()[translate(.,'FALSE','false')='false']">false</xsl:template>

  <!-- object -->
  <xsl:template match="*" name="base">
    <xsl:if test="not(preceding-sibling::*)">{</xsl:if>
    <xsl:call-template name="escape-string">
      <xsl:with-param name="s" select="name()"/>
    </xsl:call-template>
    <xsl:text>:</xsl:text>
    <!-- check type of node -->
    <xsl:choose>
      <!-- null nodes -->
      <xsl:when test="count(child::node())=0">null</xsl:when>
      <!-- other nodes -->
      <xsl:otherwise>
        <xsl:apply-templates select="child::node()"/>
      </xsl:otherwise>
    </xsl:choose>
    <!-- end of type check -->
    <xsl:if test="following-sibling::*">,</xsl:if>
    <xsl:if test="not(following-sibling::*)">}</xsl:if>
  </xsl:template>

  <!-- array -->
  <xsl:template match="*[count(../*[name(../*)=name(.)])=count(../*) and count(../*)' ||chr(38) ||'gt;1]">
    <xsl:if test="not(preceding-sibling::*)">[</xsl:if>
    <xsl:choose>
      <xsl:when test="not(child::node())">
        <xsl:text>null</xsl:text>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="child::node()"/>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="following-sibling::*">,</xsl:if>
    <xsl:if test="not(following-sibling::*)">]</xsl:if>
  </xsl:template>
  
  <!-- convert root element to an anonymous container -->
  <xsl:template match="/">
    <xsl:apply-templates select="node()"/>
  </xsl:template>
    
</xsl:stylesheet>^';
  
  End Get_Xml_To_Json_Stylesheet;

  Function Ref_Cursor_To_Json(p_Ref_Cursor In Sys_Refcursor, p_Max_Rows In Number := Null, p_Skip_Rows In Number := Null)
    Return Json_List As
    l_Ctx         Dbms_Xmlgen.Ctxhandle;
    l_Num_Rows    Pls_Integer;
    l_Xml         Xmltype;
    l_Json        Xmltype;
    l_Returnvalue Clob;
  Begin
  
    l_Ctx := Dbms_Xmlgen.Newcontext(p_Ref_Cursor);
  
    Dbms_Xmlgen.Setnullhandling(l_Ctx, Dbms_Xmlgen.Empty_Tag);
  
    -- for pagination
  
    If p_Max_Rows Is Not Null Then
      Dbms_Xmlgen.Setmaxrows(l_Ctx, p_Max_Rows);
    End If;
  
    If p_Skip_Rows Is Not Null Then
      Dbms_Xmlgen.Setskiprows(l_Ctx, p_Skip_Rows);
    End If;
  
    -- get the XML content
    l_Xml := Dbms_Xmlgen.Getxmltype(l_Ctx, Dbms_Xmlgen.None);
  
    l_Num_Rows := Dbms_Xmlgen.Getnumrowsprocessed(l_Ctx);
  
    Dbms_Xmlgen.Closecontext(l_Ctx);
  
    Close p_Ref_Cursor;
  
    If l_Num_Rows > 0 Then
      -- perform the XSL transformation
      l_Json        := l_Xml.Transform(Xmltype(Get_Xml_To_Json_Stylesheet));
      l_Returnvalue := l_Json.Getclobval();
    Else
      l_Returnvalue := g_Json_Null_Object;
    End If;
  
    l_Returnvalue := Dbms_Xmlgen.Convert(l_Returnvalue, Dbms_Xmlgen.Entity_Decode);
  
    If (l_Num_Rows = 0) Then
      Return Json_List();
    Else
      If (l_Num_Rows = 1) Then
        Declare
          Ret Json_List := Json_List();
        Begin
          Ret.Append(Json(Json(l_Returnvalue).Get('ROWSET')).Get('ROW'));
          Return Ret;
        End;
      Else
        Return Json_List(Json(l_Returnvalue).Get('ROWSET'));
      End If;
    End If;
  
  Exception
    When Scanner_Exception Then
      Dbms_Output.Put('Scanner problem with the following input: ');
      Dbms_Output.Put_Line(l_Returnvalue);
      Raise;
    When Parser_Exception Then
      Dbms_Output.Put('Parser problem with the following input: ');
      Dbms_Output.Put_Line(l_Returnvalue);
      Raise;
    When Others Then
      Raise;
  End Ref_Cursor_To_Json;

  Function Sql_To_Json(p_Sql In Varchar2, p_Max_Rows In Number := Null, p_Skip_Rows In Number := Null) Return Json_List As
    v_Cur Sys_Refcursor;
  Begin
    Open v_Cur For p_Sql;
    Return Ref_Cursor_To_Json(v_Cur, p_Max_Rows, p_Skip_Rows);
  
  End Sql_To_Json;

End Json_Util_Pkg;
/

Create Or Replace Package Json_Helper As
  Function Merge(p_a_Json Json, p_b_Json Json) Return Json;

  Function Join(p_a_List Json_List, p_b_List Json_List) Return Json_List;

  Function Keep(p_Json Json, p_Keys Json_List) Return Json;

  Function Remove(p_Json Json, p_Keys Json_List) Return Json;

  --equals
  Function Equals(p_V1 Json_Value, p_V2 Json_Value, Exact Boolean Default True) Return Boolean;
  Function Equals(p_V1 Json_Value, p_V2 Json, Exact Boolean Default True) Return Boolean;
  Function Equals(p_V1 Json_Value, p_V2 Json_List, Exact Boolean Default True) Return Boolean;
  Function Equals(p_V1 Json_Value, p_V2 Number) Return Boolean;
  Function Equals(p_V1 Json_Value, p_V2 Varchar2) Return Boolean;
  Function Equals(p_V1 Json_Value, p_V2 Boolean) Return Boolean;
  Function Equals(p_V1 Json_Value, p_V2 Clob) Return Boolean;
  Function Equals(p_V1 Json, p_V2 Json, Exact Boolean Default True) Return Boolean;
  Function Equals(p_V1 Json_List, p_V2 Json_List, Exact Boolean Default True) Return Boolean;

  --contains json, json_value
  --contains json_list, json_value
  Function Contains(p_V1 Json, p_V2 Json_Value, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json, p_V2 Json, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json, p_V2 Json_List, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json, p_V2 Number, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json, p_V2 Varchar2, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json, p_V2 Boolean, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json, p_V2 Clob, Exact Boolean Default False) Return Boolean;

  Function Contains(p_V1 Json_List, p_V2 Json_Value, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json_List, p_V2 Json, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json_List, p_V2 Json_List, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json_List, p_V2 Number, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json_List, p_V2 Varchar2, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json_List, p_V2 Boolean, Exact Boolean Default False) Return Boolean;
  Function Contains(p_V1 Json_List, p_V2 Clob, Exact Boolean Default False) Return Boolean;

End Json_Helper;
/

Create Or Replace Package Body Json_Helper As

  --recursive merge
  Function Merge(p_a_Json Json, p_b_Json Json) Return Json As
    l_Json      Json;
    l_Jv        Json_Value;
    l_Indx      Number;
    l_Recursive Json_Value;
  Begin
    -- Initialize our return object
    l_Json := p_a_Json;
  
    -- loop through p_b_json
    l_Indx := p_b_Json.Json_Data.First;
    Loop
      Exit When l_Indx Is Null;
      l_Jv := p_b_Json.Json_Data(l_Indx);
      If (l_Jv.Is_Object) Then
        --recursive
        l_Recursive := l_Json.Get(l_Jv.Mapname);
        If (l_Recursive Is Not Null And l_Recursive.Is_Object) Then
          l_Json.Put(l_Jv.Mapname, Merge(Json(l_Recursive), Json(l_Jv)));
        Else
          l_Json.Put(l_Jv.Mapname, l_Jv);
        End If;
      Else
        l_Json.Put(l_Jv.Mapname, l_Jv);
      End If;
    
      --increment
      l_Indx := p_b_Json.Json_Data.Next(l_Indx);
    End Loop;
  
    Return l_Json;
  
  End Merge;

  -- join two lists
  Function Join(p_a_List Json_List, p_b_List Json_List) Return Json_List As
    l_Json_List Json_List := p_a_List;
  Begin
    For Indx In 1 .. p_b_List.Count Loop
      l_Json_List.Append(p_b_List.Get(Indx));
    End Loop;
  
    Return l_Json_List;
  
  End Join;

  -- keep keys.
  Function Keep(p_Json Json, p_Keys Json_List) Return Json As
    l_Json  Json := Json();
    Mapname Varchar2(4000);
  Begin
    For i In 1 .. p_Keys.Count Loop
      Mapname := p_Keys.Get(i).Get_String;
      If (p_Json.Exist(Mapname)) Then
        l_Json.Put(Mapname, p_Json.Get(Mapname));
      End If;
    End Loop;
  
    Return l_Json;
  End Keep;

  -- drop keys.
  Function Remove(p_Json Json, p_Keys Json_List) Return Json As
    l_Json Json := p_Json;
  Begin
    For i In 1 .. p_Keys.Count Loop
      l_Json.Remove(p_Keys.Get(i).Get_String);
    End Loop;
  
    Return l_Json;
  End Remove;

  --equals functions

  Function Equals(p_V1 Json_Value, p_V2 Number) Return Boolean As
  Begin
    If (p_V2 Is Null) Then
      Return p_V1.Is_Null;
    End If;
  
    If (Not p_V1.Is_Number) Then
      Return False;
    End If;
  
    Return p_V2 = p_V1.Get_Number;
  End;

  Function Equals(p_V1 Json_Value, p_V2 Boolean) Return Boolean As
  Begin
    If (p_V2 Is Null) Then
      Return p_V1.Is_Null;
    End If;
  
    If (Not p_V1.Is_Bool) Then
      Return False;
    End If;
  
    Return p_V2 = p_V1.Get_Bool;
  End;

  Function Equals(p_V1 Json_Value, p_V2 Varchar2) Return Boolean As
  Begin
    If (p_V2 Is Null) Then
      Return(p_V1.Is_Null Or p_V1.Get_String Is Null);
    End If;
  
    If (Not p_V1.Is_String) Then
      Return False;
    End If;
  
    Return p_V2 = p_V1.Get_String;
  End;

  Function Equals(p_V1 Json_Value, p_V2 Clob) Return Boolean As
    My_Clob Clob;
    Res     Boolean;
  Begin
    If (p_V2 Is Null) Then
      Return p_V1.Is_Null;
    End If;
  
    If (Not p_V1.Is_String) Then
      Return False;
    End If;
  
    My_Clob := Empty_Clob();
    Dbms_Lob.Createtemporary(My_Clob, True);
    p_V1.Get_String(My_Clob);
  
    Res := Dbms_Lob.Compare(p_V2, My_Clob) = 0;
    Dbms_Lob.Freetemporary(My_Clob);
    Return Res;
  End;

  Function Equals(p_V1 Json_Value, p_V2 Json_Value, Exact Boolean) Return Boolean As
  Begin
    If (p_V2 Is Null Or p_V2.Is_Null) Then
      Return(p_V1 Is Null Or p_V1.Is_Null);
    End If;
  
    If (p_V2.Is_Number) Then
      Return Equals(p_V1, p_V2.Get_Number);
    End If;
    If (p_V2.Is_Bool) Then
      Return Equals(p_V1, p_V2.Get_Bool);
    End If;
    If (p_V2.Is_Object) Then
      Return Equals(p_V1, Json(p_V2), Exact);
    End If;
    If (p_V2.Is_Array) Then
      Return Equals(p_V1, Json_List(p_V2), Exact);
    End If;
    If (p_V2.Is_String) Then
      If (p_V2.Extended_Str Is Null) Then
        Return Equals(p_V1, p_V2.Get_String);
      Else
        Declare
          My_Clob Clob;
          Res     Boolean;
        Begin
          My_Clob := Empty_Clob();
          Dbms_Lob.Createtemporary(My_Clob, True);
          p_V2.Get_String(My_Clob);
          Res := Equals(p_V1, My_Clob);
          Dbms_Lob.Freetemporary(My_Clob);
          Return Res;
        End;
      End If;
    End If;
  
    Return False; --should never happen
  End;

  Function Equals(p_V1 Json_Value, p_V2 Json_List, Exact Boolean) Return Boolean As
    Cmp Json_List;
    Res Boolean := True;
  Begin
    --  p_v1.print(false);
    --  p_v2.print(false);
    --  dbms_output.put_line('labc1'||case when exact then 'X' else 'U' end);
  
    If (p_V2 Is Null) Then
      Return p_V1.Is_Null;
    End If;
  
    If (Not p_V1.Is_Array) Then
      Return False;
    End If;
  
    --  dbms_output.put_line('labc2'||case when exact then 'X' else 'U' end);
  
    Cmp := Json_List(p_V1);
    If (Cmp.Count != p_V2.Count And Exact) Then
      Return False;
    End If;
  
    --  dbms_output.put_line('labc3'||case when exact then 'X' else 'U' end);
  
    If (Exact) Then
      For i In 1 .. Cmp.Count Loop
        Res := Equals(Cmp.Get(i), p_V2.Get(i), Exact);
        If (Not Res) Then
          Return Res;
        End If;
      End Loop;
    Else
      --  dbms_output.put_line('labc4'||case when exact then 'X' else 'U' end);
      If (p_V2.Count > Cmp.Count) Then
        Return False;
      End If;
      --  dbms_output.put_line('labc5'||case when exact then 'X' else 'U' end);
    
      --match sublist here!
      For x In 0 .. (Cmp.Count - p_V2.Count) Loop
        --  dbms_output.put_line('labc7'||x);
      
        For i In 1 .. p_V2.Count Loop
          Res := Equals(Cmp.Get(x + i), p_V2.Get(i), Exact);
          If (Not Res) Then
            Goto Next_Index;
          End If;
        End Loop;
        Return True;
      
        <<next_Index>>
        Null;
      End Loop;
    
      --  dbms_output.put_line('labc7'||case when exact then 'X' else 'U' end);
    
      Return False; --no match
    
    End If;
  
    Return Res;
  End;

  Function Equals(p_V1 Json_Value, p_V2 Json, Exact Boolean) Return Boolean As
    Cmp Json;
    Res Boolean := True;
  Begin
    --  p_v1.print(false);
    --  p_v2.print(false);
    --  dbms_output.put_line('abc1');
  
    If (p_V2 Is Null) Then
      Return p_V1.Is_Null;
    End If;
  
    If (Not p_V1.Is_Object) Then
      Return False;
    End If;
  
    Cmp := Json(p_V1);
  
    --  dbms_output.put_line('abc2');
  
    If (Cmp.Count != p_V2.Count And Exact) Then
      Return False;
    End If;
  
    --  dbms_output.put_line('abc3');
    Declare
      K1        Json_List := p_V2.Get_Keys;
      Key_Index Number;
    Begin
      For i In 1 .. K1.Count Loop
        Key_Index := Cmp.Index_Of(K1.Get(i).Get_String);
        If (Key_Index = -1) Then
          Return False;
        End If;
        If (Exact) Then
          If (Not Equals(p_V2.Get(i), Cmp.Get(Key_Index), True)) Then
            Return False;
          End If;
        Else
          --non exact
          Declare
            V1 Json_Value := Cmp.Get(Key_Index);
            V2 Json_Value := p_V2.Get(i);
          Begin
            --  dbms_output.put_line('abc3 1/2');
            --            v1.print(false);
            --            v2.print(false);
          
            If (V1.Is_Object And V2.Is_Object) Then
              If (Not Equals(V1, V2, False)) Then
                Return False;
              End If;
            Elsif (V1.Is_Array And V2.Is_Array) Then
              If (Not Equals(V1, V2, False)) Then
                Return False;
              End If;
            Else
              If (Not Equals(V1, V2, True)) Then
                Return False;
              End If;
            End If;
          End;
        
        End If;
      End Loop;
    End;
  
    --  dbms_output.put_line('abc4');
  
    Return True;
  End;

  Function Equals(p_V1 Json, p_V2 Json, Exact Boolean) Return Boolean As
  Begin
    Return Equals(p_V1.To_Json_Value, p_V2, Exact);
  End;

  Function Equals(p_V1 Json_List, p_V2 Json_List, Exact Boolean) Return Boolean As
  Begin
    Return Equals(p_V1.To_Json_Value, p_V2, Exact);
  End;

  --contain
  Function Contains(p_V1 Json, p_V2 Json_Value, Exact Boolean) Return Boolean As
    v_Values Json_List;
  Begin
    If (Equals(p_V1.To_Json_Value, p_V2, Exact)) Then
      Return True;
    End If;
  
    v_Values := p_V1.Get_Values;
  
    For i In 1 .. v_Values.Count Loop
      Declare
        v_Val Json_Value := v_Values.Get(i);
      Begin
        If (v_Val.Is_Object) Then
          If (Contains(Json(v_Val), p_V2, Exact)) Then
            Return True;
          End If;
        End If;
        If (v_Val.Is_Array) Then
          If (Contains(Json_List(v_Val), p_V2, Exact)) Then
            Return True;
          End If;
        End If;
      
        If (Equals(v_Val, p_V2, Exact)) Then
          Return True;
        End If;
      End;
    
    End Loop;
  
    Return False;
  End;

  Function Contains(p_V1 Json_List, p_V2 Json_Value, Exact Boolean) Return Boolean As
  Begin
    If (Equals(p_V1.To_Json_Value, p_V2, Exact)) Then
      Return True;
    End If;
  
    For i In 1 .. p_V1.Count Loop
      Declare
        v_Val Json_Value := p_V1.Get(i);
      Begin
        If (v_Val.Is_Object) Then
          If (Contains(Json(v_Val), p_V2, Exact)) Then
            Return True;
          End If;
        End If;
        If (v_Val.Is_Array) Then
          If (Contains(Json_List(v_Val), p_V2, Exact)) Then
            Return True;
          End If;
        End If;
      
        If (Equals(v_Val, p_V2, Exact)) Then
          Return True;
        End If;
      End;
    
    End Loop;
  
    Return False;
  End;

  Function Contains(p_V1 Json, p_V2 Json, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, p_V2.To_Json_Value, Exact);
  End;
  Function Contains(p_V1 Json, p_V2 Json_List, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, p_V2.To_Json_Value, Exact);
  End;
  Function Contains(p_V1 Json, p_V2 Number, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;
  Function Contains(p_V1 Json, p_V2 Varchar2, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;
  Function Contains(p_V1 Json, p_V2 Boolean, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;
  Function Contains(p_V1 Json, p_V2 Clob, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;

  Function Contains(p_V1 Json_List, p_V2 Json, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, p_V2.To_Json_Value, Exact);
  End;
  Function Contains(p_V1 Json_List, p_V2 Json_List, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, p_V2.To_Json_Value, Exact);
  End;
  Function Contains(p_V1 Json_List, p_V2 Number, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;
  Function Contains(p_V1 Json_List, p_V2 Varchar2, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;
  Function Contains(p_V1 Json_List, p_V2 Boolean, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;
  Function Contains(p_V1 Json_List, p_V2 Clob, Exact Boolean) Return Boolean As
  Begin
    Return Contains(p_V1, Json_Value(p_V2), Exact);
  End;

End Json_Helper;
/

create or replace type pljson_varray as table of varchar2(32767);
/

create or replace type pljson_vtab as table of pljson_varray;
/

create or replace type pljson_narray as table of number;
/

Create Or Replace Type Pljson_Table_Impl As Object
(
  Str      Clob, -- varchar2(32767),
  Paths    Pljson_Varray,
  Names    Pljson_Varray,
  Data_Tab Pljson_Vtab,
  Row_Inds Pljson_Narray,
  Ret_Type Anytype,
  Static Function Odcitabledescribe(Rtype Out Anytype, Str Clob, Paths Pljson_Varray, Names Pljson_Varray := Null)
    Return Number,
  Static Function Odcitableprepare(Sctx Out Pljson_Table_Impl, Ti In Sys.Odcitabfuncinfo, Str Clob, Paths Pljson_Varray, Names Pljson_Varray := Null)
    Return Number,
  Static Function Odcitablestart(Sctx In Out Pljson_Table_Impl, Str Clob, Paths Pljson_Varray, Names Pljson_Varray := Null)
    Return Number,
  Member Function Odcitablefetch(Self In Out Pljson_Table_Impl, Nrows In Number, Outset Out Anydataset) Return Number,
  Member Function Odcitableclose(Self In Pljson_Table_Impl) Return Number,
  Static Function Json_Table(Str Clob, Paths Pljson_Varray, Names Pljson_Varray := Null) Return Anydataset
    Pipelined Using Pljson_Table_Impl
);
/

Create Or Replace Type Body Pljson_Table_Impl As
  Static Function Odcitabledescribe(Rtype Out Anytype, Str Clob, Paths Pljson_Varray, Names Pljson_Varray := Null)
    Return Number Is
    Atyp Anytype;
  Begin
    --dbms_output.put_line('>>Describe');
    Anytype.Begincreate(Dbms_Types.Typecode_Object, Atyp);
    If Names Is Null Then
      For i In Paths.First .. Paths.Last Loop
        Atyp.Addattr('JSON_' || Ltrim(To_Char(i)), Dbms_Types.Typecode_Varchar2, Null, Null, 32767, Null, Null);
      End Loop;
    Else
      For i In Names.First .. Names.Last Loop
        Atyp.Addattr(Upper(Names(i)), Dbms_Types.Typecode_Varchar2, Null, Null, 32767, Null, Null);
      End Loop;
    End If;
    Atyp.Endcreate;
  
    Anytype.Begincreate(Dbms_Types.Typecode_Table, Rtype);
    Rtype.Setinfo(Null, Null, Null, Null, Null, Atyp, Dbms_Types.Typecode_Object, 0);
    Rtype.Endcreate();
  
    Return Odciconst.Success;
  Exception
    When Others Then
      Return Odciconst.Error;
  End;

  Static Function Odcitableprepare(Sctx Out Pljson_Table_Impl, Ti In Sys.Odcitabfuncinfo, Str Clob, Paths Pljson_Varray, Names Pljson_Varray := Null)
    Return Number Is
    Elem_Typ Sys.Anytype;
    Prec     Pls_Integer;
    Scale    Pls_Integer;
    Len      Pls_Integer;
    Csid     Pls_Integer;
    Csfrm    Pls_Integer;
    Tc       Pls_Integer;
    Aname    Varchar2(30);
  Begin
    --dbms_output.put_line('>>Prepare');
  
    Tc   := Ti.Rettype.Getattreleminfo(1, Prec, Scale, Len, Csid, Csfrm, Elem_Typ, Aname);
    Sctx := Pljson_Table_Impl(Str, Paths, Names, Pljson_Vtab(), Pljson_Narray(), Elem_Typ);
    Return Odciconst.Success;
  End;

  Static Function Odcitablestart(Sctx In Out Pljson_Table_Impl, Str Clob, Paths Pljson_Varray, Names Pljson_Varray := Null)
    Return Number Is
    Json_Obj Json;
    Json_Val Json_Value;
    Buf      Varchar2(32767);
    --data_tab pljson_vtab := pljson_vtab();
    Json_Array  Json_List;
    Json_Elem   Json_Value;
    Value_Array Pljson_Varray := Pljson_Varray();
  Begin
    --dbms_output.put_line('>>Start');
  
    Sctx.Data_Tab.Delete;
    --dbms_output.put_line('json_str='||str);
    Json_Obj := Json(Str);
    For i In Paths.First .. Paths.Last Loop
      --dbms_output.put_line('path='||paths(i));
      Json_Val := Json_Ext.Get_Json_Value(Json_Obj, Paths(i));
      --dbms_output.put_line('type='||json_val.get_type());
      Case Json_Val.Typeval
      --when 1 then 'object';
        When 2 Then
          -- 'array';
          Json_Array := Json_List(Json_Val);
          Value_Array.Delete;
          For j In 1 .. Json_Array.Count Loop
            Json_Elem := Json_Array.Get(j);
            Case Json_Elem.Typeval
            --when 1 then 'object';
            --when 2 then -- 'array';
              When 3 Then
                -- 'string';
                Buf := Json_Elem.Get_String();
                --dbms_output.put_line('res[]='||buf);
                Value_Array.Extend();
                Value_Array(Value_Array.Last) := Buf;
              When 4 Then
                -- 'number';
                Buf := To_Char(Json_Elem.Get_Number());
                --dbms_output.put_line('res[]='||buf);
                Value_Array.Extend();
                Value_Array(Value_Array.Last) := Buf;
              When 5 Then
                -- 'bool';
                Buf := Case Json_Elem.Get_Bool()
                         When True Then
                          'true'
                         When False Then
                          'false'
                       End;
                --dbms_output.put_line('res[]='||buf);
                Value_Array.Extend();
                Value_Array(Value_Array.Last) := Buf;
              When 6 Then
                -- 'null';
                Buf := Null;
                --dbms_output.put_line('res[]='||buf);
                Value_Array.Extend();
                Value_Array(Value_Array.Last) := Buf;
              Else
                -- if object is unknown or does not exist add new element of type null
                Buf := Null;
                --dbms_output.put_line('res='||buf);
                Sctx.Data_Tab.Extend();
                Sctx.Data_Tab(Sctx.Data_Tab.Last) := Pljson_Varray(Buf);
            End Case;
          End Loop;
          Sctx.Data_Tab.Extend();
          Sctx.Data_Tab(Sctx.Data_Tab.Last) := Value_Array;
        When 3 Then
          -- 'string';
          Buf := Json_Val.Get_String();
          --dbms_output.put_line('res='||buf);
          Sctx.Data_Tab.Extend();
          Sctx.Data_Tab(Sctx.Data_Tab.Last) := Pljson_Varray(Buf);
        When 4 Then
          -- 'number';
          Buf := To_Char(Json_Val.Get_Number());
          --dbms_output.put_line('res='||buf);
          Sctx.Data_Tab.Extend();
          Sctx.Data_Tab(Sctx.Data_Tab.Last) := Pljson_Varray(Buf);
        When 5 Then
          -- 'bool';
          Buf := Case Json_Val.Get_Bool()
                   When True Then
                    'true'
                   When False Then
                    'false'
                 End;
          --dbms_output.put_line('res='||buf);
          Sctx.Data_Tab.Extend();
          Sctx.Data_Tab(Sctx.Data_Tab.Last) := Pljson_Varray(Buf);
        When 6 Then
          -- 'null';
          Buf := Null;
          --dbms_output.put_line('res='||buf);
          Sctx.Data_Tab.Extend();
          Sctx.Data_Tab(Sctx.Data_Tab.Last) := Pljson_Varray(Buf);
        Else
          -- if object is unknown or does not exist add new element of type null
          Buf := Null;
          --dbms_output.put_line('res='||buf);
          Sctx.Data_Tab.Extend();
          Sctx.Data_Tab(Sctx.Data_Tab.Last) := Pljson_Varray(Buf);
      End Case;
    End Loop;
  
    --dbms_output.put_line('initialize row indexes');
    Sctx.Row_Inds.Delete;
    --for i in data_tab.FIRST .. data_tab.LAST loop
    For i In Paths.First .. Paths.Last Loop
      Sctx.Row_Inds.Extend();
      Sctx.Row_Inds(Sctx.Row_Inds.Last) := 1;
    End Loop;
  
    Return Odciconst.Success;
  End;

  Member Function Odcitablefetch(Self In Out Pljson_Table_Impl, Nrows In Number, Outset Out Anydataset) Return Number Is
    j        Number;
    Num_Rows Number := 0;
  Begin
    --dbms_output.put_line('>>Fetch');
  
    Anydataset.Begincreate(Dbms_Types.Typecode_Object, Self.Ret_Type, Outset);
  
    /* iterative cartesian product algorithm */
    <<main_Loop>>
    While True Loop
      Exit When Num_Rows = Nrows Or Row_Inds(1) = 0;
      --data_row.delete;
      Outset.Addinstance;
      Outset.Piecewise();
      --dbms_output.put_line('put one row piece');
      For i In Data_Tab.First .. Data_Tab.Last Loop
        Outset.Setvarchar2(Data_Tab(i) (Row_Inds(i)));
      End Loop;
      --pipe row(data_row);
      Num_Rows := Num_Rows + 1;
    
      --dbms_output.put_line('adjust row indexes');
      j := Row_Inds.Count;
      <<index_Loop>>
      While True Loop
        Row_Inds(j) := Row_Inds(j) + 1;
        If Row_Inds(j) <= Data_Tab(j).Count Then
          Exit Index_Loop;
        End If;
        Row_Inds(j) := 1;
        j := j - 1;
        If j < 1 Then
          Row_Inds(1) := 0; -- hack to indicate end of all fetches
          Exit Main_Loop;
        End If;
      End Loop;
    End Loop;
  
    Outset.Endcreate;
  
    Return Odciconst.Success;
  End;

  Member Function Odcitableclose(Self In Pljson_Table_Impl) Return Number Is
  Begin
    --dbms_output.put_line('>>Close');
    Return Odciconst.Success;
  End;

End;
/
--1.调整了JSON_Table的授权
--2、增加了对Json_Ac的授权
--3、该授权缺失对PLJSON_NARRAY、PLJSON_VARRAY、PLJSON_VTAB的授权
--types
Grant Execute On Json To Public;
Create Or Replace Public Synonym Json For Json;
Grant Execute On Json_List To Public;
Create Or Replace Public Synonym Json_List For Json_List;
Grant Execute On Json_Value To Public;
Create Or Replace Public Synonym Json_Value For Json_Value;
Grant Execute On Json_Value_Array To Public;
Create Or Replace Public Synonym Json_Value_Array For Json_Value_Array;
Create Synonym Pljson_Table For Pljson_Table_Impl;
Create Synonym Json_Table For Pljson_Table_Impl;
Grant Execute On Json_Table To Public;
Create Or Replace Public Synonym Json_Table For Json_Table;
--packages
Grant Execute On Json_Ext To Public;
Create Or Replace Public Synonym Json_Ext For Json_Ext;
Grant Execute On Json_Parser To Public;
Create Or Replace Public Synonym Json_Parser For Json_Parser;
Grant Execute On Json_Printer To Public;
Create Or Replace Public Synonym Json_Printer For Json_Printer;
Grant Execute On Json_Ac To Public;
Create Or Replace Public Synonym Json_Ac For Json_Ac;
Grant Execute On Json_Dyn To Public;
Create Or Replace Public Synonym Json_Dyn For Json_Dyn;
Grant Execute On Json_Ml To Public;
Create Or Replace Public Synonym Json_Ml For Json_Ml;
Grant Execute On Json_Xml To Public;
Create Or Replace Public Synonym Json_Xml For Json_Xml;
Grant Execute On Json_Util_Pkg To Public;
Create Or Replace Public Synonym Json_Util_Pkg For Json_Util_Pkg;
Grant Execute On Json_Helper To Public;
Create Or Replace Public Synonym Json_Helper For Json_Helper;