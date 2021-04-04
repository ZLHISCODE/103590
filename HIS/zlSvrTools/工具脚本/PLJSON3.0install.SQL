--[版本号]3.0.X000
--版本号用于判断自我更新
--lshuo注释
--基于标准的PLJSON 3.0修改而成.根据当前Oracle版本进行不同的源码控制。
--  Oracle版本为10时，PLJSON_LIST支持字符数组或数字数组的构造函数。但是executeList无法返回字符数组或数字数组列
--  Oracle版本>=11、executeList可以返回字符数组或数字数组列。PLJSON_LIST支持字符数组或数字数组的构造函数。
--1、适用于Oracle 10g版本。根据不同Oracle 版本条件编译为不同的源码(条件编译常量10g支持)
--2、版本号为3.0.X000,X表示修正，后面数字为修正的次数。
--3、该版本可以与低于3.0版本的PLJSON并存,但是卸载该版本会导致低于3.0版本的PLJSON无法使用（可以通过卸载后重新安装解决）。卸载低版本的PLJSON不会影响该版本使用。

Declare
--11G及以上版本
--PLSQL_OPTIMIZE_LEVEL=3:PL/SQL编译器会在合适时执行自动的子程序内联,使用被调用子程序的真实副本来替换子程序调用语句以提高执行效率
--plsql_code_type='native':编译为native code（C代码），在运行时不需要进行解释，运行更快
Begin
  If Dbms_Db_Version.Version > 10 Then
    Execute Immediate 'ALTER SESSION SET PLSQL_OPTIMIZE_LEVEL = 3';
    Execute Immediate 'ALTER SESSION SET plsql_code_type = ''NATIVE''';    
  End If;
End;
/

create or replace type  pljson_element as object 
(
  obj_type number
)
not final;
/
create or replace type pljson_value  as object (
  typeval number(1), -- 1 = object, 2 = array, 3 = string, 4 = number, 5 = bool, 6 = null 
--  * Private variable for internal processing. 
  str varchar2(32767),
--  * Private variable for internal processing. 
  num number, -- store 1 as true, 0 as false 
--  * Private variable for internal processing. 
  num_double binary_double, -- both num and num_double are set, there is never exception (until Oracle 12c)
--  * Private variable for internal processing. 
  num_repr_number_p varchar2(1),
--  * Private variable for internal processing. 
  num_repr_double_p varchar2(1),
--  * Private variable for internal processing. 
  object_or_array pljson_element, -- object or array in here 
--  * Private variable for internal processing. 
  extended_str clob,

--   mapping 
--  * Private variable for internal processing. 
  mapname varchar2(4000),
--  * Private variable for internal processing. 
  mapindx number(32),

  constructor function pljson_value(elem pljson_element) return self as result,
  constructor function pljson_value(str varchar2, esc boolean default true) return self as result,
  constructor function pljson_value(str clob, esc boolean default true) return self as result,
  constructor function pljson_value(num number) return self as result,
--   E.I.Sarmas (github.com/dsnz)   2016-11-03   support for binary_double numbers 
  constructor function pljson_value(num_double binary_double) return self as result,
  constructor function pljson_value(b boolean) return self as result,
  constructor function pljson_value return self as result,

  member function get_element return pljson_element,

  static function makenull return pljson_value,

  member function get_type return varchar2,

  member function get_string(max_byte_size number default null, max_char_size number default null) return varchar2,

  member procedure get_string(self in pljson_value, buf in out nocopy clob),

  member function get_clob return clob,

  member function get_number return number,

  member function get_double return binary_double,

  member function get_bool return boolean,

  member function get_null return varchar2,

  member function is_object return boolean,

  member function is_array return boolean,

  member function is_string return boolean,

  member function is_number return boolean,

  member function is_bool return boolean,

  member function is_null return boolean,

  member function is_number_repr_number return boolean,
  member function is_number_repr_double return boolean,

  member procedure parse_number(str varchar2),

  -- this procedure is meant to be used internally only
  member function number_toString return varchar2,

--   Output methods 
  member function to_char(spaces boolean default true, chars_per_line number default 0) return varchar2,
  member procedure to_clob(self in pljson_value, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true),
  member procedure print(self in pljson_value, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null), --32512 is maximum
  member procedure htp(self in pljson_value, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null),

  member function value_of(self in pljson_value, max_byte_size number default null, max_char_size number default null) return varchar2
) not final;
/

create or replace type pljson_value_array as table of pljson_value;
/
--lshuo注释
--1.0存在会出错
begin execute immediate 'drop type pljson_varray force'; exception when others then null; end;
/
create or replace type pljson_varray as table of varchar2(32767);
/
--lshuo注释
--1.0存在会出错
begin execute immediate 'drop type pljson_narray force'; exception when others then null; end;
/
create or replace type pljson_narray as table of number;
/

create or replace type pljson_list  under pljson_element (
  list_data pljson_value_array,

  constructor function pljson_list return self as result,

  constructor function pljson_list(str varchar2) return self as result,

  constructor function pljson_list(str clob) return self as result,

  constructor function pljson_list(str blob, charset varchar2 default 'UTF8') return self as result,
  --lshuo注释
  --由于10g不支持DESC_TAB3 ，且无法返回数组对象，pljson_narray，pljson_varray
	constructor function pljson_list(str_array pljson_varray) return self as result,

	constructor function pljson_list(num_array pljson_narray) return self as result,
  constructor function pljson_list(elem pljson_value) return self as result,
  
--   list management 
  member procedure append(self in out nocopy pljson_list, elem pljson_value, position pls_integer default null),
  member procedure append(self in out nocopy pljson_list, elem varchar2, position pls_integer default null),
  member procedure append(self in out nocopy pljson_list, elem clob, position pls_integer default null),
  member procedure append(self in out nocopy pljson_list, elem number, position pls_integer default null),
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure append(self in out nocopy pljson_list, elem binary_double, position pls_integer default null),
  member procedure append(self in out nocopy pljson_list, elem boolean, position pls_integer default null),
  member procedure append(self in out nocopy pljson_list, elem pljson_list, position pls_integer default null),
  
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem pljson_value),
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem varchar2),
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem clob),
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem number),
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem binary_double),
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem boolean),
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem pljson_list),
  
  member procedure remove(self in out nocopy pljson_list, position pls_integer),
  member procedure remove_first(self in out nocopy pljson_list),
  member procedure remove_last(self in out nocopy pljson_list),
  
  member function count return number,
  member function get(position pls_integer) return pljson_value,
  member function get_string(position pls_integer) return varchar2,
  member function get_clob(position pls_integer) return clob,
  member function get_number(position pls_integer) return number,
  member function get_double(position pls_integer) return binary_double,
  member function get_bool(position pls_integer) return boolean,
  member function get_pljson_list(position pls_integer) return pljson_list,
  member function head return pljson_value,
  member function last return pljson_value,
  member function tail return pljson_list,
  
  member function to_json_value return pljson_value,
  
--   output methods 
  member function to_char(spaces boolean default true, chars_per_line number default 0) return varchar2,
  member procedure to_clob(self in pljson_list, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true),
  member procedure print(self in pljson_list, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null), --32512 is maximum
  member procedure htp(self in pljson_list, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null),
  
--   json path 
  member function path(json_path varchar2, base number default 1) return pljson_value,
--   json path_put 
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem pljson_value, base number default 1),
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem varchar2, base number default 1),
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem clob, base number default 1),
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem number, base number default 1),
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem binary_double, base number default 1),
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem boolean, base number default 1),
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem pljson_list, base number default 1),
  
--   json path_remove 
  member procedure path_remove(self in out nocopy pljson_list, json_path varchar2, base number default 1)
  
--   --backwards compatibility
--  member procedure add_elem(self in out nocopy json_list, elem json_value, position pls_integer default null),
--  member procedure add_elem(self in out nocopy json_list, elem varchar2, position pls_integer default null),
--  member procedure add_elem(self in out nocopy json_list, elem number, position pls_integer default null),
--  member procedure add_elem(self in out nocopy json_list, elem boolean, position pls_integer default null),
--  member procedure add_elem(self in out nocopy json_list, elem json_list, position pls_integer default null),
--
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem json_value),
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem varchar2),
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem number),
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem boolean),
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem json_list),
--
--  member procedure remove_elem(self in out nocopy json_list, position pls_integer),
--  member function get_elem(position pls_integer) return json_value,
--  member function get_first return json_value,
--  member function get_last return json_value
) not final;
/

create or replace type pljson  under pljson_element (
  json_data pljson_value_array,
--  * Private variable for internal processing. 
  check_for_duplicate number,

  constructor function pljson return self as result,

  constructor function pljson(str varchar2) return self as result,
 
  constructor function pljson(str in clob) return self as result,
  
  constructor function pljson(str in blob, charset varchar2 default 'UTF8') return self as result,
 
  constructor function pljson(str_array pljson_varray) return self as result,
 
  constructor function pljson(elem pljson_value) return self as result,
 
  constructor function pljson(l in out nocopy pljson_list) return self as result,
 
  member procedure remove(pair_name varchar2),
  
  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value pljson_value, position pls_integer default null),

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value varchar2, position pls_integer default null),

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value clob, position pls_integer default null),

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value number, position pls_integer default null),

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value binary_double, position pls_integer default null),

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value boolean, position pls_integer default null),

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value pljson, position pls_integer default null),
  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value pljson_list, position pls_integer default null),

  member function count return number,

  member function get(pair_name varchar2) return pljson_value,

  member function get_string(pair_name varchar2) return varchar2,
  member function get_clob(pair_name varchar2) return clob,
  member function get_number(pair_name varchar2) return number,
  member function get_double(pair_name varchar2) return binary_double,
  member function get_bool(pair_name varchar2) return boolean,
  member function get_pljson(pair_name varchar2) return pljson,
  member function get_pljson_list(pair_name varchar2) return pljson_list,

  member function get(position pls_integer) return pljson_value,

  member function index_of(pair_name varchar2) return number,

  member function exist(pair_name varchar2) return boolean,

  member function to_json_value return pljson_value,

  member procedure check_duplicate(self in out nocopy pljson, v_set boolean),
  member procedure remove_duplicates(self in out nocopy pljson),

  member function to_char(spaces boolean default true, chars_per_line number default 0) return varchar2,

  member procedure to_clob(self in pljson, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true),

  member procedure print(self in pljson, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null), --32512 is maximum

  member procedure htp(self in pljson, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null),

  member function path(json_path varchar2, base number default 1) return pljson_value,

--   json path_put 
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem pljson_value, base number default 1),
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem varchar2, base number default 1),
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem clob, base number default 1),
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem number, base number default 1),
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem binary_double, base number default 1),
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem boolean, base number default 1),
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem pljson, base number default 1),
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem pljson_list, base number default 1),

--   json path_remove 
  member procedure path_remove(self in out nocopy pljson, json_path varchar2, base number default 1),
 
  member function get_keys return pljson_list,
 
  member function get_values return pljson_list

) not final;
/
create or replace package pljson_ext as
  function parsePath(json_path varchar2, base number default 1) return pljson_list;

  --JSON Path getters
  function get_json_value(obj pljson, v_path varchar2, base number default 1) return pljson_value;
  function get_string(obj pljson, path varchar2,       base number default 1) return varchar2;
  function get_number(obj pljson, path varchar2,       base number default 1) return number;
  function get_double(obj pljson, path varchar2,       base number default 1) return binary_double;
  function get_json(obj pljson, path varchar2,         base number default 1) return pljson;
  function get_json_list(obj pljson, path varchar2,    base number default 1) return pljson_list;
  function get_bool(obj pljson, path varchar2,         base number default 1) return boolean;

  --JSON Path putters
  procedure put(obj in out nocopy pljson, path varchar2, elem varchar2,   base number default 1);
  procedure put(obj in out nocopy pljson, path varchar2, elem number,     base number default 1);
  procedure put(obj in out nocopy pljson, path varchar2, elem binary_double, base number default 1);
  procedure put(obj in out nocopy pljson, path varchar2, elem pljson,       base number default 1);
  procedure put(obj in out nocopy pljson, path varchar2, elem pljson_list,  base number default 1);
  procedure put(obj in out nocopy pljson, path varchar2, elem boolean,    base number default 1);
  procedure put(obj in out nocopy pljson, path varchar2, elem pljson_value, base number default 1);

  procedure remove(obj in out nocopy pljson, path varchar2, base number default 1);

  --Pretty print with JSON Path - obsolete in 0.9.4 - obj.path(v_path).(to_char,print,htp)
  function pp(obj pljson, v_path varchar2) return varchar2;
  procedure pp(obj pljson, v_path varchar2); --using dbms_output.put_line
  procedure pp_htp(obj pljson, v_path varchar2); --using htp.print

  --extra function checks if number has no fraction
  function is_integer(v pljson_value) return boolean;

  format_string varchar2(30 char) := 'yyyy-mm-dd hh24:mi:ss';
  --extension enables json to store dates without compromising the implementation
  function to_json_value(d date) return pljson_value;
  --notice that a date type in json is also a varchar2
  function is_date(v pljson_value) return boolean;
  --conversion is needed to extract dates
  function to_date(v pljson_value) return date;
  
  function to_date2(v pljson_value) return date;
  --JSON Path with date
  function get_date(obj pljson, path varchar2, base number default 1) return date;
  procedure put(obj in out nocopy pljson, path varchar2, elem date, base number default 1);

  function encodeBase64Blob2Clob(p_blob in  blob) return clob;
  function decodeBase64Clob2Blob(p_clob clob) return blob;

  function base64(binarydata blob) return pljson_list;
  function base64(l pljson_list) return blob;

  function encode(binarydata blob) return pljson_value;
  function decode(v pljson_value) return blob;

--  
--    implemented as a procedure to force you to declare the CLOB so you can free it later
--  
  procedure blob2clob(b blob, c out clob, charset varchar2 default 'UTF8');
end pljson_ext;
/

create or replace package pljson_parser as

  type rToken IS RECORD (
    type_name VARCHAR2(7),
    line PLS_INTEGER,
    col PLS_INTEGER,
    data VARCHAR2(32767),
    data_overflow clob); -- max_string_size

  type lTokens is table of rToken index by pls_integer;
  type json_src is record (len number, offset number, src varchar2(32767), s_clob clob);

  json_strict boolean not null := false;

  function next_char(indx number, s in out nocopy json_src) return varchar2;
  function next_char2(indx number, s in out nocopy json_src, amount number default 1) return varchar2;
  function parseObj(tokens lTokens, indx in out nocopy pls_integer) return pljson;

  function prepareClob(buf in clob) return pljson_parser.json_src;
  function prepareVarchar2(buf in varchar2) return pljson_parser.json_src;
  function lexer(jsrc in out nocopy json_src) return lTokens;
  procedure print_token(t rToken);

  function parser(str varchar2) return pljson;
  function parse_list(str varchar2) return pljson_list;
  function parse_any(str varchar2) return pljson_value;
  function parser(str clob) return pljson;
  function parse_list(str clob) return pljson_list;
  function parse_any(str clob) return pljson_value;
  procedure remove_duplicates(obj in out nocopy pljson);
  function get_version return varchar2;

end pljson_parser;
/

create or replace package body pljson_parser as
  decimalpoint varchar2(1 char) := '.';

  procedure update_decimalpoint as
  begin
--    select substr(value, 1, 1)
--    into decimalpoint
--    from nls_session_parameters
--    where parameter = 'NLS_NUMERIC_CHARACTERS';
--by:zyk,小数点分隔符用缺省的，避免查询数据库以提高性能
    null;
  end update_decimalpoint;
 
  function next_char(indx number, s in out nocopy json_src) return varchar2 as
  begin
    if (indx > s.len) then return null; end if;
    --right offset 
--     if (indx > 4000 + s.offset or indx < s.offset) then 
--     fix for issue #37 
    if (indx > 4000 + s.offset or indx <= s.offset) then
    --load right offset
      s.offset := indx - (indx mod 4000);
--       addon fix for issue #37 
      if s.offset = indx then
        s.offset := s.offset - 4000;
      end if;
      s.src := dbms_lob.substr(s.s_clob, 4000, s.offset+1);
    end if;
    --read from s.src
    return substr(s.src, indx-s.offset, 1);
  end;

  function next_char2(indx number, s in out nocopy json_src, amount number default 1) return varchar2 as
    buf varchar2(32767) := '';
  begin
    for i in 1..amount loop
      buf := buf || next_char(indx-1+i, s);
    end loop;
    return buf;
  end;

  function prepareClob(buf clob) return pljson_parser.json_src as
    temp pljson_parser.json_src;
  begin
    temp.s_clob := buf;
    temp.offset := 0;
    --by:zyk,优化代码，少读一次变量
    --temp.src := dbms_lob.substr(buf, 4000, temp.offset+1);
    temp.src := dbms_lob.substr(buf, 4000, 1);
    temp.len := dbms_lob.getlength(buf);
    return temp;
  end;

  function prepareVarchar2(buf varchar2) return pljson_parser.json_src as
    temp pljson_parser.json_src;
  begin
    temp.s_clob := buf;
    temp.offset := 0;
    temp.src := substr(buf, 1, 4000);
    temp.len := length(buf);
    return temp;
  end;

  procedure debug(text varchar2) as
  begin
    dbms_output.put_line(text);
  end;

  procedure print_token(t rToken) as
  begin
    dbms_output.put_line('Line: '||t.line||' - Column: '||t.col||' - Type: '||t.type_name||' - Content: '||t.data);
  end print_token;

--   SCANNER FUNCTIONS START 
  procedure s_error(text varchar2, line number, col number) as
  begin
    raise_application_error(-20100, 'JSON Scanner exception @ line: '||line||' column: '||col||' - '||text);
  end;

  procedure s_error(text varchar2, tok rToken) as
  begin
    raise_application_error(-20100, 'JSON Scanner exception @ line: '||tok.line||' column: '||tok.col||' - '||text);
  end;

  function mt(t varchar2, l pls_integer, c pls_integer, d varchar2) return rToken as
    token rToken;
  begin
    token.type_name := t;
    token.line := l;
    token.col := c;
    token.data := d;
    return token;
  end;

  function lexNumber(jsrc in out nocopy json_src, tok in out nocopy rToken, indx in out nocopy pls_integer) return pls_integer as
    numbuf varchar2(4000) := '';
    buf varchar2(4);
    checkLoop boolean;
    indx_begin pls_integer;
    offset_begin pls_integer;
  begin
    buf := next_char(indx, jsrc);

    --by:zyk,在indx没有变的情况下，之前的代码重复读了一次next_char
    --if (buf = '-') then numbuf := '-'; indx := indx + 1; end if;
    --buf := next_char(indx, jsrc);
    if (buf = '-') then 
        numbuf := '-'; indx := indx + 1; 
        buf := next_char(indx, jsrc);
    end if;

    --0 or [1-9]([0-9])*
    if (buf = '0') then
      numbuf := numbuf || '0'; indx := indx + 1;
      buf := next_char(indx, jsrc);
    elsif (buf >= '1' and buf <= '9') then
      numbuf := numbuf || buf; indx := indx + 1;
      --read digits
      buf := next_char(indx, jsrc);

      --by:zyk,一次性取出以优化效率(减少字符拼接次数)
      indx_begin:=indx;offset_begin:=jsrc.offset;

      while (buf >= '0' and buf <= '9') loop        
        indx := indx + 1;
        buf := next_char(indx, jsrc);
        IF offset_begin<>jsrc.offset THEN
            offset_begin:=jsrc.offset;

            indx:=indx-1;
            buf := next_char(indx, jsrc);
            numbuf := numbuf ||substr(jsrc.src,indx_begin-jsrc.offset,indx+1-indx_begin);

            indx_begin:=indx+1;
        END if;        
      end loop;
      IF indx_begin<>indx THEN numbuf := numbuf ||substr(jsrc.src,indx_begin-jsrc.offset,indx-indx_begin); END IF;
    end if;
    --fraction
    if (buf = '.') then
      numbuf := numbuf || buf; indx := indx + 1;
      buf := next_char(indx, jsrc);

      --by:zyk,一次性取出以优化效率(减少字符拼接次数)
      indx_begin:=indx;offset_begin:=jsrc.offset;

      checkLoop := FALSE;      
      while (buf >= '0' and buf <= '9') loop
        checkLoop := TRUE;
        indx := indx + 1;
        buf := next_char(indx, jsrc);
        IF offset_begin<>jsrc.offset THEN
            offset_begin:=jsrc.offset;

            indx:=indx-1;
            buf := next_char(indx, jsrc);
            numbuf := numbuf ||substr(jsrc.src,indx_begin-jsrc.offset,indx+1-indx_begin);             
            
            indx_begin:=indx+1;
        END if;
      end loop;
      if (not checkLoop) then
        s_error('Expected: digits in fraction', tok);
      ELSE 
        numbuf := numbuf ||substr(jsrc.src,indx_begin-jsrc.offset,indx-indx_begin);
      end if;
    end if;
    --exp part
    if (buf in ('e', 'E')) then
      numbuf := numbuf || buf; indx := indx + 1;
      buf := next_char(indx, jsrc);
      if (buf = '+' or buf = '-') then
        numbuf := numbuf || buf; indx := indx + 1;
        buf := next_char(indx, jsrc);
      end if;
      checkLoop := FALSE;
      while (buf >= '0' and buf <= '9') loop
        checkLoop := TRUE;
        numbuf := numbuf || buf; indx := indx + 1;
        buf := next_char(indx, jsrc);
      end loop;
      if (not checkLoop) then
        s_error('Expected: digits in exp', tok);
      end if;
    end if;

    tok.data := numbuf;
    return indx;
  end lexNumber;

  -- [a-zA-Z]([a-zA-Z0-9])
  function lexName(jsrc in out nocopy json_src, tok in out nocopy rToken, indx in out nocopy pls_integer) return pls_integer as
    varbuf varchar2(32767) := '';
    buf varchar(4);
    num number;
  begin
    buf := next_char(indx, jsrc);
    while (REGEXP_LIKE(buf, '^[[:alnum:]\_]$', 'i')) loop
      varbuf := varbuf || buf;
      indx := indx + 1;
      buf := next_char(indx, jsrc);
      if (buf is null) then
        goto retname;
        --debug('Premature string ending');
      end if;
    end loop;
    <<retname>>
    --could check for reserved keywords here
    --debug(varbuf);
    tok.data := varbuf;
    return indx-1;
  end lexName;

  procedure updateClob(v_extended in out nocopy clob, v_str varchar2) as
  begin
    dbms_lob.writeappend(v_extended, length(v_str), v_str);
  end updateClob;

  function lexString(jsrc in out nocopy json_src, tok in out nocopy rToken, indx in out nocopy pls_integer, endChar char) return pls_integer as
    v_extended clob := null; v_count number := 0;
    varbuf varchar2(32767) := '';--by:zyk,虽然超过8192该值就会重置，但分片拼接时会超出
    buf varchar(4);
    wrong boolean;
    v_node varchar2(8000);--虽然一个片段(src)只有4000个字符，但是可能有汉字(每个字符占两个字节)，所以要定义为8000
    n_end boolean;
  begin
    indx := indx +1;
    buf := next_char(indx, jsrc);

    while (buf != endChar) loop
      --clob control
      if (v_count > 8191) then --crazy oracle error (16383 is the highest working length with unistr - 8192 choosen to be safe)
        if (v_extended is null) then
          v_extended := empty_clob();
          dbms_lob.createtemporary(v_extended, true);
        end if;
        updateClob(v_extended, unistr(varbuf));
        varbuf := ''; v_count := 0;
      end if;
      if (buf = Chr(13) or buf = CHR(9) or buf = CHR(10)) then
        s_error('Control characters not allowed (CHR(9),CHR(10),CHR(13))', tok);
      end if;
      if (buf = '\') then
        --varbuf := varbuf || buf;
        indx := indx + 1;
        buf := next_char(indx, jsrc);
        case
          when buf in ('\') then
            varbuf := varbuf || buf || buf; v_count := v_count + 2;
            indx := indx + 1;
            buf := next_char(indx, jsrc);
          when buf in ('"', '/') then
            varbuf := varbuf || buf; v_count := v_count + 1;
            indx := indx + 1;
            buf := next_char(indx, jsrc);
          when buf = '''' then
            if (json_strict = false) then
              varbuf := varbuf || buf; v_count := v_count + 1;
              indx := indx + 1;
              buf := next_char(indx, jsrc);
            else
              s_error('strictmode - expected: " \ / b f n r t u ', tok);
            end if;
          when buf in ('b', 'f', 'n', 'r', 't') then
            --backspace b = U+0008
            --formfeed  f = U+000C
            --newline   n = U+000A
            --carret    r = U+000D
            --tabulator t = U+0009
            case buf
            when 'b' then varbuf := varbuf || chr(8);
            when 'f' then varbuf := varbuf || chr(12);
            when 'n' then varbuf := varbuf || chr(10);
            when 'r' then varbuf := varbuf || chr(13);
            when 't' then varbuf := varbuf || chr(9);
            end case;
            --varbuf := varbuf || buf;
            v_count := v_count + 1;
            indx := indx + 1;
            buf := next_char(indx, jsrc);
          when buf = 'u' then
            --four hexadecimal chars
            declare
              four varchar2(4);
            begin
              four := next_char2(indx+1, jsrc, 4);
              wrong := FALSE;
              if (upper(substr(four, 1, 1)) not in ('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','a','b','c','d','e','f')) then wrong := TRUE; end if;
              if (upper(substr(four, 2, 1)) not in ('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','a','b','c','d','e','f')) then wrong := TRUE; end if;
              if (upper(substr(four, 3, 1)) not in ('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','a','b','c','d','e','f')) then wrong := TRUE; end if;
              if (upper(substr(four, 4, 1)) not in ('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','a','b','c','d','e','f')) then wrong := TRUE; end if;
              if (wrong) then
                s_error('expected: " \u([0-9][A-F]){4}', tok);
              end if;
--              varbuf := varbuf || buf || four;
              varbuf := varbuf || '\'||four;--chr(to_number(four,'XXXX'));
              v_count := v_count + 5;
              indx := indx + 5;
              buf := next_char(indx, jsrc);
              end;
          else
            s_error('expected: " \ / b f n r t u ', tok);
        end case;
      else
        varbuf := varbuf || buf; v_count := v_count + 1;
        indx := indx + 1;
        --by:zyk,一次性取出以优化效率
        --可能跨片段(每片段4000),没有找到结束符则把剩余的全取
        if instr(jsrc.src,endChar,indx-jsrc.offset)=0 then
          v_node:=substr(jsrc.src,indx-jsrc.offset);
          n_end:=false;
        else
           v_node:=substr(jsrc.src,indx-jsrc.offset,instr(jsrc.src,endChar,indx-jsrc.offset)-(indx-jsrc.offset));
           n_end:=true;
        end if;
        --字符串的值为""，也会为空
        if v_node is not null And instr(v_node,'\')=0 then
            varbuf := varbuf || v_node;
            v_count:= Length(varbuf);
            if v_count>8191 then
              if (v_extended is null) then
                v_extended := empty_clob();
                dbms_lob.createtemporary(v_extended, true);
              end if;
              updateClob(v_extended, unistr(varbuf));
              varbuf := '';
              v_count:= 0;
            end if;

            indx := indx + length(v_node);
            if n_end then
              buf := endChar;
              exit;
            else
              buf := next_char(indx, jsrc);
            end if;
        else
            buf := next_char(indx, jsrc);
        end if;
      end if;
    end loop;

    if (buf is null) then
      s_error('string ending not found', tok);
      --debug('Premature string ending');
    end if;

    --debug(varbuf);
    --dbms_output.put_line(varbuf);
    if (v_extended is not null) then
      if varbuf is not null then
         updateClob(v_extended, unistr(varbuf));
      end if;
      tok.data_overflow := v_extended;
      tok.data := dbms_lob.substr(v_extended, 1, 32767);
    else
	  -- unistr(varbuf) 生僻字会报错
      tok.data := unistr(asciistr(varbuf));
    end if;
    return indx;
  end lexString;

  function lexer(jsrc in out nocopy json_src) return lTokens as
    tokens lTokens;
    indx pls_integer := 1;
    tok_indx pls_integer := 1;
    buf varchar2(4);
    lin_no number := 1;
    col_no number := 0;
  begin
    while (indx <= jsrc.len) loop
      --read into buf
      buf := next_char(indx, jsrc);
      col_no := col_no + 1;
      --convert to switch case
      case
        when buf = '{' then tokens(tok_indx) := mt('{', lin_no, col_no, null); tok_indx := tok_indx + 1;
        when buf = '}' then tokens(tok_indx) := mt('}', lin_no, col_no, null); tok_indx := tok_indx + 1;
        when buf = ',' then tokens(tok_indx) := mt(',', lin_no, col_no, null); tok_indx := tok_indx + 1;
        when buf = ':' then tokens(tok_indx) := mt(':', lin_no, col_no, null); tok_indx := tok_indx + 1;
        when (buf in ('-', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')) then --number
          tokens(tok_indx) := mt('NUMBER', lin_no, col_no, null);
          indx := lexNumber(jsrc, tokens(tok_indx), indx)-1;
          col_no := col_no + length(tokens(tok_indx).data);
          tok_indx := tok_indx + 1;
        when buf = '"' then --number
          tokens(tok_indx) := mt('STRING', lin_no, col_no, null);
          indx := lexString(jsrc, tokens(tok_indx), indx, '"');
          col_no := col_no + length(tokens(tok_indx).data) + 1;
          tok_indx := tok_indx + 1;
        when buf = '[' then tokens(tok_indx) := mt('[', lin_no, col_no, null); tok_indx := tok_indx + 1;
        when buf = ']' then tokens(tok_indx) := mt(']', lin_no, col_no, null); tok_indx := tok_indx + 1;
        --by:zyk,Begin,调整顺序，减少高频内容的判断次数
        when buf = '''' and json_strict = false then --number
          tokens(tok_indx) := mt('STRING', lin_no, col_no, null);
          indx := lexString(jsrc, tokens(tok_indx), indx, '''');
          col_no := col_no + length(tokens(tok_indx).data) + 1; --hovsa her
          tok_indx := tok_indx + 1;
        when buf = 't' then
          if (next_char2(indx, jsrc, 4) != 'true') then
            if (json_strict = false and REGEXP_LIKE(buf, '^[[:alpha:]]$', 'i')) then
              tokens(tok_indx) := mt('STRING', lin_no, col_no, null);
              indx := lexName(jsrc, tokens(tok_indx), indx);
              col_no := col_no + length(tokens(tok_indx).data) + 1;
              tok_indx := tok_indx + 1;
            else
              s_error('Expected: ''true''', lin_no, col_no);
            end if;
          else
            tokens(tok_indx) := mt('TRUE', lin_no, col_no, null); tok_indx := tok_indx + 1;
            indx := indx + 3;
            col_no := col_no + 3;
          end if;
        when buf = 'n' then
          if (next_char2(indx, jsrc, 4) != 'null') then
            if (json_strict = false and REGEXP_LIKE(buf, '^[[:alpha:]]$', 'i')) then
              tokens(tok_indx) := mt('STRING', lin_no, col_no, null);
              indx := lexName(jsrc, tokens(tok_indx), indx);
              col_no := col_no + length(tokens(tok_indx).data) + 1;
              tok_indx := tok_indx + 1;
            else
              s_error('Expected: ''null''', lin_no, col_no);
            end if;
          else
            tokens(tok_indx) := mt('NULL', lin_no, col_no, null); tok_indx := tok_indx + 1;
            indx := indx + 3;
            col_no := col_no + 3;
          end if;
        when buf = 'f' then
          if (next_char2(indx, jsrc, 5) != 'false') then
            if (json_strict = false and REGEXP_LIKE(buf, '^[[:alpha:]]$', 'i')) then
              tokens(tok_indx) := mt('STRING', lin_no, col_no, null);
              indx := lexName(jsrc, tokens(tok_indx), indx);
              col_no := col_no + length(tokens(tok_indx).data) + 1;
              tok_indx := tok_indx + 1;
            else
              s_error('Expected: ''false''', lin_no, col_no);
            end if;
          else
            tokens(tok_indx) := mt('FALSE', lin_no, col_no, null); tok_indx := tok_indx + 1;
            indx := indx + 4;
            col_no := col_no + 4;
          end if;
--         -- 9 = TAB, 10 = \n, 13 = \r (Linux = \n, Windows = \r\n, Mac = \r 
        when (buf = Chr(10)) then --linux newlines
          lin_no := lin_no + 1;
          col_no := 0;

        when (buf = Chr(13)) then --Windows or Mac way
          lin_no := lin_no + 1;
          col_no := 0;
          if (jsrc.len >= indx +1) then -- better safe than sorry
            buf := next_char(indx+1, jsrc);
            if (buf = Chr(10)) then --\r\n
              indx := indx + 1;
            end if;
          end if;

        when (buf = CHR(9)) then null; --tabbing
        --by:zyk,end(alpha的判断必须放到tnf等判断之后)
        when json_strict = false and REGEXP_LIKE(buf, '^[[:alpha:]]$', 'i') then
          tokens(tok_indx) := mt('STRING', lin_no, col_no, null);
          indx := lexName(jsrc, tokens(tok_indx), indx);
          if (tokens(tok_indx).data_overflow is not null) then
            col_no := col_no + dbms_lob.getlength(tokens(tok_indx).data_overflow) + 1;
          else
            col_no := col_no + length(tokens(tok_indx).data) + 1;
          end if;
          tok_indx := tok_indx + 1;
        when json_strict = false and buf||next_char(indx+1, jsrc) = '/*' then --strip comments
          declare
            saveindx number := indx;
            un_esc clob;
          begin
            indx := indx + 1;
            loop
              indx := indx + 1;
              buf := next_char(indx, jsrc)||next_char(indx+1, jsrc);
              exit when buf = '*/';
              exit when buf is null;
            end loop;

            if (indx = saveindx+2) then
              --enter unescaped mode              
              un_esc := empty_clob();
              dbms_lob.createtemporary(un_esc, true);
              indx := indx + 1;
              loop
                indx := indx + 1;
                buf := next_char(indx, jsrc)||next_char(indx+1, jsrc)||next_char(indx+2, jsrc)||next_char(indx+3, jsrc);
                exit when buf = '/**/';
                if buf is null then
                  s_error('Unexpected sequence /**/ to end unescaped data: '||buf, lin_no, col_no);
                end if;
                buf := next_char(indx, jsrc);
                dbms_lob.writeappend(un_esc, length(buf), buf);
              end loop;
              tokens(tok_indx) := mt('ESTRING', lin_no, col_no, null);
              tokens(tok_indx).data_overflow := un_esc;
              col_no := col_no + dbms_lob.getlength(un_esc) + 1; 
              tok_indx := tok_indx + 1;
              indx := indx + 2;
            end if;

            indx := indx + 1;
          end;
        when buf = ' ' then null;  
	--lshuo注释
	--部分VB传递的参数，可能包含了Chr(0)，导致失败
        --when buf = CHR(0) then null;  
        else
          s_error('Unexpected char: '||buf, lin_no, col_no);
      end case;

      indx := indx + 1;
    end loop;

    return tokens;
  end lexer;

--   SCANNER END 

--   PARSER FUNCTIONS START 
  procedure p_error(text varchar2, tok rToken) as
  begin
    raise_application_error(-20101, 'JSON Parser exception @ line: '||tok.line||' column: '||tok.col||' - '||text);
  end;

  function parseArr(tokens lTokens, indx in out nocopy pls_integer) return pljson_list as
    e_arr pljson_value_array := pljson_value_array();
    ret_list pljson_list := pljson_list();
    v_count number := 0;
    tok rToken;
    pv pljson_value;
  begin
    --value, value, value ]
    if (indx > tokens.count) then p_error('more elements in array was excepted', tok); end if;
    tok := tokens(indx);
    while (tok.type_name != ']') loop
      e_arr.extend;
      v_count := v_count + 1;
      case tok.type_name
        when 'TRUE' then e_arr(v_count) := pljson_value(true);
        when 'FALSE' then e_arr(v_count) := pljson_value(false);
        when 'NULL' then e_arr(v_count) := pljson_value();
        when 'STRING' then e_arr(v_count) := case when tok.data_overflow is not null then pljson_value(tok.data_overflow) else pljson_value(tok.data) end;
        when 'ESTRING' then e_arr(v_count) := pljson_value(tok.data_overflow, false);
--         E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
        
        when 'NUMBER' then
          pv := pljson_value(0);
          pv.parse_number(replace(tok.data, '.', decimalpoint));
          e_arr(v_count) := pv;
        when '[' then
          declare e_list pljson_list; begin
            indx := indx + 1;
            e_list := parseArr(tokens, indx);
            e_arr(v_count) := e_list.to_json_value;
          end;
        when '{' then
          indx := indx + 1;
          e_arr(v_count) := parseObj(tokens, indx).to_json_value;
        else
          p_error('Expected a value', tok);
      end case;
      indx := indx + 1;
      if (indx > tokens.count) then p_error('] not found', tok); end if;
      tok := tokens(indx);
      if (tok.type_name = ',') then --advance
        indx := indx + 1;
        if (indx > tokens.count) then p_error('more elements in array was excepted', tok); end if;
        tok := tokens(indx);
        if (tok.type_name = ']') then --premature exit
          p_error('Premature exit in array', tok);
        end if;
      elsif (tok.type_name != ']') then --error
        p_error('Expected , or ]', tok);
      end if;

    end loop;
    ret_list.list_data := e_arr;
    return ret_list;
  end parseArr;

  function parseMem(tokens lTokens, indx in out pls_integer, mem_name varchar2, mem_indx number) return pljson_value as
    mem pljson_value;
    tok rToken;
    pv pljson_value;
  begin
    tok := tokens(indx);
    case tok.type_name
      when 'TRUE' then mem := pljson_value(true);
      when 'FALSE' then mem := pljson_value(false);
      when 'NULL' then mem := pljson_value();
      when 'STRING' then mem := case when tok.data_overflow is not null then pljson_value(tok.data_overflow) else pljson_value(tok.data) end;
      when 'ESTRING' then mem := pljson_value(tok.data_overflow, false);
--       E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
      
      when 'NUMBER' then
        pv := pljson_value(0);
        pv.parse_number(replace(tok.data, '.', decimalpoint));
        mem := pv;
      when '[' then
        declare
          e_list pljson_list;
        begin
          indx := indx + 1;
          e_list := parseArr(tokens, indx);
          mem := e_list.to_json_value;
        end;
      when '{' then
        indx := indx + 1;
        mem := parseObj(tokens, indx).to_json_value;
      else
        p_error('Found '||tok.type_name, tok);
    end case;
    mem.mapname := mem_name;
    mem.mapindx := mem_indx;

    indx := indx + 1;
    return mem;
  end parseMem;

  function parseObj(tokens lTokens, indx in out nocopy pls_integer) return pljson as
    type memmap is table of number index by varchar2(4000);
    mymap memmap;
    nullelemfound boolean := false;

    obj pljson;
    tok rToken;
    mem_name varchar(4000);
    arr pljson_value_array := pljson_value_array();
  begin
    --what to expect 
    while (indx <= tokens.count) loop
      tok := tokens(indx);
      --debug('E: '||tok.type_name);
      case tok.type_name
      when 'STRING' then
        --member
        mem_name := substr(tok.data, 1, 4000);
        begin
          if (mem_name is null) then
            if (nullelemfound) then
              p_error('Duplicate empty member: ', tok);
            else
              nullelemfound := true;
            end if;
          elsif (mymap(mem_name) is not null) then
            p_error('Duplicate member name: '||mem_name, tok);
          end if;
        exception
          when no_data_found then mymap(mem_name) := 1;
        end;

        indx := indx + 1;
        if (indx > tokens.count) then p_error('Unexpected end of input', tok); end if;
        tok := tokens(indx);
        indx := indx + 1;
        if (indx > tokens.count) then p_error('Unexpected end of input', tok); end if;
        if (tok.type_name = ':') then
          --parse
          declare
            jmb pljson_value;
            x number;
          begin
            x := arr.count + 1;
            jmb := parseMem(tokens, indx, mem_name, x);
            arr.extend;
            arr(x) := jmb;
          end;
        else
          p_error('Expected '':''', tok);
        end if;
        
        if (indx > tokens.count) then p_error('Unexpected end of input', tok); end if;

        tok := tokens(indx);
        if (tok.type_name = ',') then
          --debug('found ,');
          indx := indx + 1;
          tok := tokens(indx);
          if (tok.type_name = '}') then --premature exit
            p_error('Premature exit in json object', tok);
          end if;
        elsif (tok.type_name != '}') then
           p_error('A comma seperator is probably missing', tok);
        end if;
      when '}' then
        obj := pljson();
        obj.json_data := arr;
        return obj;
      else
        p_error('Expected string or }', tok);
      end case;
    end loop;

    p_error('} not found', tokens(indx-1));

    return obj;

  end;

  function parser(str varchar2) return pljson as
    tokens lTokens;
    obj pljson;
    indx pls_integer := 1;
    jsrc json_src;
  begin
    update_decimalpoint();
    jsrc := prepareVarchar2(str);
    tokens := lexer(jsrc);
    if (tokens(indx).type_name = '{') then
      indx := indx + 1;
      obj := parseObj(tokens, indx);
    else
      raise_application_error(-20101, 'JSON Parser exception - no { start found');
    end if;
    if (tokens.count != indx) then
      p_error('} should end the JSON object', tokens(indx));
    end if;

    return obj;
  end parser;

  function parse_list(str varchar2) return pljson_list as
    tokens lTokens;
    obj pljson_list;
    indx pls_integer := 1;
    jsrc json_src;
  begin
    update_decimalpoint();
    jsrc := prepareVarchar2(str);
    tokens := lexer(jsrc);
    if (tokens(indx).type_name = '[') then
      indx := indx + 1;
      obj := parseArr(tokens, indx);
    else
      raise_application_error(-20101, 'JSON List Parser exception - no [ start found');
    end if;
    if (tokens.count != indx) then
      p_error('] should end the JSON List object', tokens(indx));
    end if;

    return obj;
  end parse_list;

  function parse_list(str clob) return pljson_list as
    tokens lTokens;
    obj pljson_list;
    indx pls_integer := 1;
    jsrc json_src;
  begin
    update_decimalpoint();
    jsrc := prepareClob(str);
    tokens := lexer(jsrc);
    if (tokens(indx).type_name = '[') then
      indx := indx + 1;
      obj := parseArr(tokens, indx);
    else
      raise_application_error(-20101, 'JSON List Parser exception - no [ start found');
    end if;
    if (tokens.count != indx) then
      p_error('] should end the JSON List object', tokens(indx));
    end if;

    return obj;
  end parse_list;

  function parser(str clob) return pljson as
    tokens lTokens;
    obj pljson;
    indx pls_integer := 1;
    jsrc json_src;
  begin
    update_decimalpoint();
    --dbms_output.put_line('Using clob');
    jsrc := prepareClob(str);
    tokens := lexer(jsrc);
    if (tokens(indx).type_name = '{') then
      indx := indx + 1;
      obj := parseObj(tokens, indx);
    else
      raise_application_error(-20101, 'JSON Parser exception - no { start found');
    end if;
    if (tokens.count != indx) then
      p_error('} should end the JSON object', tokens(indx));
    end if;

    return obj;
  end parser;

  function parse_any(str varchar2) return pljson_value as
    tokens lTokens;
    obj pljson_list;
    ret pljson_value;
    indx pls_integer := 1;
    jsrc json_src;
  begin
    update_decimalpoint();
    jsrc := prepareVarchar2(str);
    tokens := lexer(jsrc);
    tokens(tokens.count+1).type_name := ']';
    obj := parseArr(tokens, indx);
    if (tokens.count != indx) then
      p_error('] should end the JSON List object', tokens(indx));
    end if;

    return obj.head();
  end parse_any;

  function parse_any(str clob) return pljson_value as
    tokens lTokens;
    obj pljson_list;
    indx pls_integer := 1;
    jsrc json_src;
  begin
    update_decimalpoint();
    jsrc := prepareClob(str);
    tokens := lexer(jsrc);
    tokens(tokens.count+1).type_name := ']';
    obj := parseArr(tokens, indx);
    if (tokens.count != indx) then
      p_error('] should end the JSON List object', tokens(indx));
    end if;

    return obj.head();
  end parse_any;

--   last entry is the one to keep 
  procedure remove_duplicates(obj in out nocopy pljson) as
    type memberlist is table of pljson_value index by varchar2(4000);
    members memberlist;
    nulljsonvalue pljson_value := null;
    validated pljson := pljson();
    indx varchar2(4000);
  begin
    for i in 1 .. obj.count loop
      if (obj.get(i).mapname is null) then
        nulljsonvalue := obj.get(i);
      else
        members(obj.get(i).mapname) := obj.get(i);
      end if;
    end loop;

    validated.check_duplicate(false);
    indx := members.first;
    loop
      exit when indx is null;
      validated.put(indx, members(indx));
      indx := members.next(indx);
    end loop;
    if (nulljsonvalue is not null) then
      validated.put('', nulljsonvalue);
    end if;

    validated.check_for_duplicate := obj.check_for_duplicate;

    obj := validated;
  end;

  function get_version return varchar2 as
  begin
	return 'PL/JSON v3.0.x001';
  end get_version;

end pljson_parser;
/

create or replace package pljson_printer as
  indent_string varchar2(10 char) := '  '; --chr(9); for tab
  newline_char varchar2(2 char)   := chr(13)||chr(10); -- Windows style
  --newline_char varchar2(2) := chr(10); -- Mac style
  --newline_char varchar2(2) := chr(13); -- Linux style
  --lshuo注释
  --缺省多字节转义(\uXXXX)，修改为多字节不转义
  ascii_output boolean    not null := false;
  empty_string_as_null boolean not null := false;
  escape_solidus boolean  not null := false;
  
  function pretty_print(obj pljson, spaces boolean default true, line_length number default 0) return varchar2;
  function pretty_print_list(obj pljson_list, spaces boolean default true, line_length number default 0) return varchar2;
  function pretty_print_any(json_part pljson_value, spaces boolean default true, line_length number default 0) return varchar2;
  procedure pretty_print(obj pljson, spaces boolean default true, buf in out nocopy clob, line_length number default 0, erase_clob boolean default true);
  procedure pretty_print_list(obj pljson_list, spaces boolean default true, buf in out nocopy clob, line_length number default 0, erase_clob boolean default true);
  procedure pretty_print_any(json_part pljson_value, spaces boolean default true, buf in out nocopy clob, line_length number default 0, erase_clob boolean default true);
  
  procedure dbms_output_clob(my_clob clob, delim varchar2, jsonp varchar2 default null);
  procedure htp_output_clob(my_clob clob, jsonp varchar2 default null);
  -- made public just for testing/profiling...
  function escapeString(str varchar2) return varchar2;

end pljson_printer;
/

create or replace package body pljson_printer as
  max_line_len number := 0;
  cur_line_len number := 0;
  
  -- associative array used inside escapeString to cache the escaped version of every character
  -- escaped so far  (example: char_map('"') contains the  '\"' string)
  -- (if the character does not need to be escaped, the character is stored unchanged in the array itself)
  -- type Rmap_char is record(buf varchar2(40), len integer);
  type Tmap_char_string is table of varchar2(40) index by varchar2(1 char); -- index by unicode char 
  char_map Tmap_char_string;
  -- since char_map the associative array is a global variable reused across multiple calls to escapeString,
  -- i need to be able to detect that the escape_solidus or ascii_output global parameters have been changed,
  -- in order to clear it and avoid using escape sequences that have been cached using the previous values
  char_map_escape_solidus boolean := escape_solidus;
  char_map_ascii_output boolean := ascii_output;
  
  function llcheck(str in varchar2) return varchar2 as
  begin
    --dbms_output.put_line(cur_line_len || ' : ' || str);
    if (max_line_len > 0 and length(str)+cur_line_len > max_line_len) then
      cur_line_len := length(str);
      return newline_char || str;
    else
      cur_line_len := cur_line_len + length(str);
      return str;
    end if;
  end llcheck;
  
  -- escapes a single character.
  function escapeChar(ch char) return varchar2 deterministic is
     result varchar2(20);
  begin
      --backspace b = U+0008
      --formfeed  f = U+000C
      --newline   n = U+000A
      --carret    r = U+000D
      --tabulator t = U+0009
      result := ch;
      
      case ch
      when chr( 8) then result := '\b';
      when chr( 9) then result := '\t';
      when chr(10) then result := '\n';
      when chr(12) then result := '\f';
      when chr(13) then result := '\r';
      when chr(34) then result := '\"';
      when chr(47) then if (escape_solidus) then result := '\/'; end if;
      when chr(92) then result := '\\';
      else if (ascii(ch) < 32) then
             result :=  '\u' || replace(substr(to_char(ascii(ch), 'XXXX'), 2, 4), ' ', '0');
        elsif (ascii_output) then
             result := replace(asciistr(ch), '\', '\u');
        end if;
      end case;
      return result;
  end;
  
  function escapeString(str varchar2) return varchar2 as
    sb varchar2(32767 byte) := '';
    buf varchar2(40);
    ch varchar2(1 char); -- unicode char 
  begin
    if (str is null) then return ''; end if;
    
    -- clear the cache if global parameters have been changed
    if char_map_escape_solidus <> escape_solidus or
       char_map_ascii_output   <> ascii_output
    then
       char_map.delete;
       char_map_escape_solidus := escape_solidus;
       char_map_ascii_output := ascii_output;
    end if;
    
    for i in 1 .. length(str) loop
      ch := substr(str, i, 1 ) ;
      
      begin
         -- it this char has already been processed, I have cached its escaped value
         buf:=char_map(ch);
      exception when no_Data_found then
         -- otherwise, i convert the value and add it to the cache
         buf := escapeChar(ch);
         char_map(ch) := buf;
      end;
      
      sb := sb || buf;
    end loop;
    return sb;
  end escapeString;
  
  function newline(spaces boolean) return varchar2 as
  begin
    cur_line_len := 0;
    if (spaces) then return newline_char; else return ''; end if;
  end;
  
--  function get_schema return varchar2 as
--  begin
--    return sys_context('userenv', 'current_schema');
--  end;
--
  function tab(indent number, spaces boolean) return varchar2 as
    i varchar(200) := '';
  begin
    if (not spaces) then return ''; end if;
    for x in 1 .. indent loop i := i || indent_string; end loop;
    return i;
  end;
  
  function getCommaSep(spaces boolean) return varchar2 as
  begin
    if (spaces) then return ', '; else return ','; end if;
  end;
  
  function getMemName(mem pljson_value, spaces boolean) return varchar2 as
  begin
    if (spaces) then
      return llcheck('"'||escapeString(mem.mapname)||'"') || llcheck(' : ');
    else
      return llcheck('"'||escapeString(mem.mapname)||'"') || llcheck(':');
    end if;
  end;
  
--   Clob method start here 
  procedure add_to_clob(buf_lob in out nocopy clob, buf_str in out nocopy varchar2, str varchar2) as
  begin
    if (lengthb(str) > 32767 - lengthb(buf_str)) then
--      dbms_lob.append(buf_lob, buf_str);
      dbms_lob.writeappend(buf_lob, length(buf_str), buf_str);
      buf_str := str;
    else
      buf_str := buf_str || str;
    end if;
  end add_to_clob;
  
  procedure flush_clob(buf_lob in out nocopy clob, buf_str in out nocopy varchar2) as
  begin
    --dbms_lob.append(buf_lob, buf_str);
    dbms_lob.writeappend(buf_lob, length(buf_str), buf_str);
  end flush_clob;
  
  procedure ppObj(obj pljson, indent number, buf in out nocopy clob, spaces boolean, buf_str in out nocopy varchar2);
  
  procedure ppString(elem pljson_value, buf in out nocopy clob, buf_str in out nocopy varchar2) is
    offset number := 1;
--     E.I.Sarmas (github.com/dsnz)   2016-01-21   limit to 5000 chars 
    v_str varchar(5000 char);
    amount number := 5000; -- chunk size for use in escapeString; maximum escaped unicode string size for chunk may be 6 one-byte chars * 5000 chunk size in multi-byte chars = 30000 1-byte chars (maximum value is 32767 1-byte chars) 
  begin
    if empty_string_as_null and elem.extended_str is null and elem.str is null then
      add_to_clob(buf, buf_str, 'null');
    else
      add_to_clob(buf, buf_str, case when elem.num = 1 then '"' else '/**/' end);
      if (elem.extended_str is not null) then --clob implementation
        while (offset <= dbms_lob.getlength(elem.extended_str)) loop
          dbms_lob.read(elem.extended_str, amount, offset, v_str);
          if (elem.num = 1) then
            add_to_clob(buf, buf_str, escapeString(v_str));
          else
            add_to_clob(buf, buf_str, v_str);
          end if;
          offset := offset + amount;
        end loop;
      else
        if (elem.num = 1) then
          while (offset <= length(elem.str)) loop
            v_str:=substr(elem.str, offset, amount);
            add_to_clob(buf, buf_str, escapeString(v_str));
            offset := offset + amount;
          end loop;
        else
          add_to_clob(buf, buf_str, elem.str);
        end if;
      end if;
      add_to_clob(buf, buf_str, case when elem.num = 1 then '"' else '/**/' end);
    end if;
  end;
  
  procedure ppEA(input pljson_list, indent number, buf in out nocopy clob, spaces boolean, buf_str in out nocopy varchar2) as
    elem pljson_value;
    arr pljson_value_array := input.list_data;
    numbuf varchar2(4000);
  begin
    for y in 1 .. arr.count loop
      elem := arr(y);
      if (elem is not null) then
      case elem.typeval
--         number 
        when 4 then
          numbuf := elem.number_toString();
          add_to_clob(buf, buf_str, llcheck(numbuf));
--         string 
        when 3 then
          ppString(elem, buf, buf_str);
--         bool 
        when 5 then
          if (elem.get_bool()) then
            add_to_clob(buf, buf_str, llcheck('true'));
          else
            add_to_clob(buf, buf_str, llcheck('false'));
          end if;
--         null 
        when 6 then
          add_to_clob(buf, buf_str, llcheck('null'));
--         array 
        when 2 then
          add_to_clob(buf, buf_str, llcheck('['));
          ppEA(pljson_list(elem), indent, buf, spaces, buf_str);
          add_to_clob(buf, buf_str, llcheck(']'));
--         object 
        when 1 then
          ppObj(pljson(elem), indent, buf, spaces, buf_str);
        else
          add_to_clob(buf, buf_str, llcheck(elem.get_type));
      end case;
      end if;
      if (y != arr.count) then add_to_clob(buf, buf_str, llcheck(getCommaSep(spaces))); end if;
    end loop;
  end ppEA;
  
  procedure ppMem(mem pljson_value, indent number, buf in out nocopy clob, spaces boolean, buf_str in out nocopy varchar2) as
    numbuf varchar2(4000);
  begin
    add_to_clob(buf, buf_str, llcheck(tab(indent, spaces)) || llcheck(getMemName(mem, spaces)));
    case mem.typeval
--       number 
      when 4 then
        numbuf := mem.number_toString();
        add_to_clob(buf, buf_str, llcheck(numbuf));
--       string 
      when 3 then
        ppString(mem, buf, buf_str);
--       bool 
      when 5 then
        if (mem.get_bool()) then
          add_to_clob(buf, buf_str, llcheck('true'));
        else
          add_to_clob(buf, buf_str, llcheck('false'));
        end if;
--       null 
      when 6 then
        add_to_clob(buf, buf_str, llcheck('null'));
--       array 
      when 2 then
        add_to_clob(buf, buf_str, llcheck('['));
        ppEA(pljson_list(mem), indent, buf, spaces, buf_str);
        add_to_clob(buf, buf_str, llcheck(']'));
--       object 
      when 1 then
        ppObj(pljson(mem), indent, buf, spaces, buf_str);
      else
        add_to_clob(buf, buf_str, llcheck(mem.get_type));
    end case;
  end ppMem;
  
  procedure ppObj(obj pljson, indent number, buf in out nocopy clob, spaces boolean, buf_str in out nocopy varchar2) as
  begin
    add_to_clob(buf, buf_str, llcheck('{') || newline(spaces));
    for m in 1 .. obj.json_data.count loop
      ppMem(obj.json_data(m), indent+1, buf, spaces, buf_str);
      if (m != obj.json_data.count) then
        add_to_clob(buf, buf_str, llcheck(',') || newline(spaces));
      else
        add_to_clob(buf, buf_str, newline(spaces));
      end if;
    end loop;
    add_to_clob(buf, buf_str, llcheck(tab(indent, spaces)) || llcheck('}')); -- || chr(13);
  end ppObj;
  
  procedure pretty_print(obj pljson, spaces boolean default true, buf in out nocopy clob, line_length number default 0, erase_clob boolean default true) as
    buf_str varchar2(32767);
    amount number := dbms_lob.getlength(buf);
  begin
    if (erase_clob and amount > 0) then dbms_lob.trim(buf, 0); dbms_lob.erase(buf, amount); end if;
    
    max_line_len := line_length;
    cur_line_len := 0;
    ppObj(obj, 0, buf, spaces, buf_str);
    flush_clob(buf, buf_str);
  end;
  
  procedure pretty_print_list(obj pljson_list, spaces boolean default true, buf in out nocopy clob, line_length number default 0, erase_clob boolean default true) as
    buf_str varchar2(32767);
    amount number := dbms_lob.getlength(buf);
  begin
    if (erase_clob and amount > 0) then dbms_lob.trim(buf, 0); dbms_lob.erase(buf, amount); end if;
    
    max_line_len := line_length;
    cur_line_len := 0;
    add_to_clob(buf, buf_str, llcheck('['));
    ppEA(obj, 0, buf, spaces, buf_str);
    add_to_clob(buf, buf_str, llcheck(']'));
    flush_clob(buf, buf_str);
  end;
  
  procedure pretty_print_any(json_part pljson_value, spaces boolean default true, buf in out nocopy clob, line_length number default 0, erase_clob boolean default true) as
    buf_str varchar2(32767) := '';
    numbuf varchar2(4000);
    amount number := dbms_lob.getlength(buf);
  begin
    if (erase_clob and amount > 0) then dbms_lob.trim(buf, 0); dbms_lob.erase(buf, amount); end if;
    
    case json_part.typeval
--       number 
      when 4 then
        numbuf := json_part.number_toString();
        add_to_clob(buf, buf_str, numbuf);
--       string 
      when 3 then
        ppString(json_part, buf, buf_str);
--       bool 
      when 5 then
        if (json_part.get_bool()) then
          add_to_clob(buf, buf_str, 'true');
        else
          add_to_clob(buf, buf_str, 'false');
        end if;
--       null 
      when 6 then
        add_to_clob(buf, buf_str, 'null');
--       array 
      when 2 then
        pretty_print_list(pljson_list(json_part), spaces, buf, line_length);
        return;
--       object 
      when 1 then
        pretty_print(pljson(json_part), spaces, buf, line_length);
        return;
      else
        add_to_clob(buf, buf_str, 'unknown type:' || json_part.get_type);
    end case;
    flush_clob(buf, buf_str);
  end;
  
--   Clob method end here 
  
--   Varchar2 method start here 
  procedure add_buf (buf in out nocopy varchar2, str in varchar2) as
  begin
    if (lengthb(str)>32767-lengthb(buf)) then
      raise_application_error(-20001,'Length of result JSON more than 32767 bytes. Use to_clob() procedures');
    end if;
    buf := buf || str;
  end;
  
  procedure ppString(elem pljson_value, buf in out nocopy varchar2) is
    offset number := 1;
--     E.I.Sarmas (github.com/dsnz)   2016-01-21   limit to 5000 chars 
    v_str varchar(5000 char);
    amount number := 5000; -- chunk size for use in escapeString; maximum escaped unicode string size for chunk may be 6 one-byte chars * 5000 chunk size in multi-byte chars = 30000 1-byte chars (maximum value is 32767 1-byte chars) 
  begin
    if empty_string_as_null and elem.extended_str is null and elem.str is null then
      add_buf(buf, 'null');
    else
      add_buf(buf, case when elem.num = 1 then '"' else '/**/' end);
      if (elem.extended_str is not null) then --clob implementation
        while (offset <= dbms_lob.getlength(elem.extended_str)) loop
          dbms_lob.read(elem.extended_str, amount, offset, v_str);
          if (elem.num = 1) then
            add_buf(buf, escapeString(v_str));
          else
            add_buf(buf, v_str);
          end if;
          offset := offset + amount;
        end loop;
      else
        if (elem.num = 1) then
          while (offset <= length(elem.str)) loop
            v_str:=substr(elem.str, offset, amount);
            add_buf(buf, escapeString(v_str));
            offset := offset + amount;
          end loop;
        else
          add_buf(buf, elem.str);
        end if;
      end if;
      add_buf(buf, case when elem.num = 1 then '"' else '/**/' end);
    end if;
  end;
  
  procedure ppObj(obj pljson, indent number, buf in out nocopy varchar2, spaces boolean);
  
  procedure ppEA(input pljson_list, indent number, buf in out varchar2, spaces boolean) as
    elem pljson_value;
    arr pljson_value_array := input.list_data;
    str varchar2(400);
  begin
    for y in 1 .. arr.count loop
      elem := arr(y);
      if (elem is not null) then
      case elem.typeval
--         number 
        when 4 then
          str := elem.number_toString();
          add_buf(buf, llcheck(str));
--         string 
        when 3 then
          ppString(elem, buf);
--         bool 
        when 5 then
          if (elem.get_bool()) then
            add_buf (buf, llcheck('true'));
          else
            add_buf (buf, llcheck('false'));
          end if;
--         null 
        when 6 then
          add_buf (buf, llcheck('null'));
--         array 
        when 2 then
          add_buf( buf, llcheck('['));
          ppEA(pljson_list(elem), indent, buf, spaces);
          add_buf( buf, llcheck(']'));
--         object 
        when 1 then
          ppObj(pljson(elem), indent, buf, spaces);
        else
          add_buf (buf, llcheck(elem.get_type)); -- should never happen 
      end case;
      end if;
      if (y != arr.count) then add_buf(buf, llcheck(getCommaSep(spaces))); end if;
    end loop;
  end ppEA;
  
  procedure ppMem(mem pljson_value, indent number, buf in out nocopy varchar2, spaces boolean) as
    str varchar2(400) := '';
  begin
    add_buf(buf, llcheck(tab(indent, spaces)) || getMemName(mem, spaces));
    case mem.typeval
--       number 
      when 4 then
        str := mem.number_toString();
        add_buf(buf, llcheck(str));
--       string 
      when 3 then
        ppString(mem, buf);
--       bool 
      when 5 then
        if (mem.get_bool()) then
          add_buf(buf, llcheck('true'));
        else
          add_buf(buf, llcheck('false'));
        end if;
--       null 
      when 6 then
        add_buf(buf, llcheck('null'));
--       array 
      when 2 then
        add_buf(buf, llcheck('['));
        ppEA(pljson_list(mem), indent, buf, spaces);
        add_buf(buf, llcheck(']'));
--       object 
      when 1 then
        ppObj(pljson(mem), indent, buf, spaces);
      else
        add_buf(buf, llcheck(mem.get_type)); -- should never happen 
    end case;
  end ppMem;
  
  procedure ppObj(obj pljson, indent number, buf in out nocopy varchar2, spaces boolean) as
  begin
    add_buf (buf, llcheck('{') || newline(spaces));
    for m in 1 .. obj.json_data.count loop
      ppMem(obj.json_data(m), indent+1, buf, spaces);
      if (m != obj.json_data.count) then
        add_buf(buf, llcheck(',') || newline(spaces));
      else
        add_buf(buf, newline(spaces));
      end if;
    end loop;
    add_buf(buf, llcheck(tab(indent, spaces)) || llcheck('}')); -- || chr(13);
  end ppObj;
  
  function pretty_print(obj pljson, spaces boolean default true, line_length number default 0) return varchar2 as
    buf varchar2(32767 byte) := '';
  begin
    max_line_len := line_length;
    cur_line_len := 0;
    ppObj(obj, 0, buf, spaces);
    return buf;
  end pretty_print;
  
  function pretty_print_list(obj pljson_list, spaces boolean default true, line_length number default 0) return varchar2 as
    buf varchar2(32767 byte) :='';
  begin
    max_line_len := line_length;
    cur_line_len := 0;
    add_buf(buf, llcheck('['));
    ppEA(obj, 0, buf, spaces);
    add_buf(buf, llcheck(']'));
    return buf;
  end;
  
  function pretty_print_any(json_part pljson_value, spaces boolean default true, line_length number default 0) return varchar2 as
    buf varchar2(32767) := '';
  begin
    case json_part.typeval
--       number 
      when 4 then
        buf := json_part.number_toString();
--       string 
      when 3 then
        ppString(json_part, buf);
--       bool 
      when 5 then
        if (json_part.get_bool()) then buf := 'true'; else buf := 'false'; end if;
--       null 
      when 6 then
        buf := 'null';
--       array 
      when 2 then
        buf := pretty_print_list(pljson_list(json_part), spaces, line_length);
--       object 
      when 1 then
        buf := pretty_print(pljson(json_part), spaces, line_length);
      else
        buf := 'weird error: ' || json_part.get_type;
    end case;
    return buf;
  end;
  
  procedure dbms_output_clob(my_clob clob, delim varchar2, jsonp varchar2 default null) as
    prev number := 1;
    indx number := 1;
    size_of_nl number := lengthb(delim);
    v_str varchar2(32767);
    amount number := 8191; -- max unicode chars 
  begin
    if (jsonp is not null) then dbms_output.put_line(jsonp||'('); end if;
    while (indx != 0) loop
      --read every line
      indx := dbms_lob.instr(my_clob, delim, prev+1);
      --dbms_output.put_line(prev || ' to ' || indx);
      
      if (indx = 0) then
        --emit from prev to end;
        amount := 8191; -- max unicode chars 
        --dbms_output.put_line(' mycloblen ' || dbms_lob.getlength(my_clob));
        loop
          dbms_lob.read(my_clob, amount, prev, v_str);
          dbms_output.put_line(v_str);
          prev := prev+amount-1;
          exit when prev >= dbms_lob.getlength(my_clob);
        end loop;
      else
        amount := indx - prev;
        if (amount > 8191) then -- max unicode chars 
          amount := 8191; -- max unicode chars 
          --dbms_output.put_line(' mycloblen ' || dbms_lob.getlength(my_clob));
          loop
            dbms_lob.read(my_clob, amount, prev, v_str);
            dbms_output.put_line(v_str);
            prev := prev+amount-1;
            amount := indx - prev;
            exit when prev >= indx - 1;
            if (amount > 8191) then amount := 8191; end if; -- max unicode chars 
          end loop;
          prev := indx + size_of_nl;
        else
          dbms_lob.read(my_clob, amount, prev, v_str);
          dbms_output.put_line(v_str);
          prev := indx + size_of_nl;
        end if;
      end if;
    
    end loop;
    if (jsonp is not null) then dbms_output.put_line(')'); end if;
    
--    while (amount != 0) loop
--      indx := dbms_lob.instr(my_clob, delim, prev+1);
--
----      dbms_output.put_line(prev || ' to ' || indx);
--      if (indx = 0) then
--        indx := dbms_lob.getlength(my_clob)+1;
--      end if;
--      if (indx-prev > 32767) then
--        indx := prev+32767;
--      end if;
----      dbms_output.put_line(prev || ' to ' || indx);
--      --substr doesnt work properly on all platforms! (come on oracle - error on Oracle VM for virtualbox)
----        dbms_output.put_line(dbms_lob.substr(my_clob, indx-prev, prev));
--      amount := indx-prev;
----        dbms_output.put_line('amount'||amount);
--      dbms_lob.read(my_clob, amount, prev, v_str);
--      dbms_output.put_line(v_str);
--      prev := indx+size_of_nl;
--      if (amount = 32767) then prev := prev-size_of_nl-1; end if;
--    end loop;
--    if (jsonp is not null) then dbms_output.put_line(')'); end if;
  end;
  
--  procedure dbms_output_clob(my_clob clob, delim varchar2, jsonp varchar2 default null) as
--    prev number := 1;
--    indx number := 1;
--    size_of_nl number := lengthb(delim);
--    v_str varchar2(32767);
--    amount number;
--  begin
--    if (jsonp is not null) then dbms_output.put_line(jsonp||'('); end if;
--    while (indx != 0) loop
--      indx := dbms_lob.instr(my_clob, delim, prev+1);
--
--      --dbms_output.put_line(prev || ' to ' || indx);
--      if (indx-prev > 32767) then
--        indx := prev+32767;
--      end if;
--      --dbms_output.put_line(prev || ' to ' || indx);
--      --substr doesnt work properly on all platforms! (come on oracle - error on Oracle VM for virtualbox)
--      if (indx = 0) then
--        --dbms_output.put_line(dbms_lob.substr(my_clob, dbms_lob.getlength(my_clob)-prev+size_of_nl, prev));
--        amount := dbms_lob.getlength(my_clob)-prev+size_of_nl;
--        dbms_lob.read(my_clob, amount, prev, v_str);
--      else
--        --dbms_output.put_line(dbms_lob.substr(my_clob, indx-prev, prev));
--        amount := indx-prev;
--        --dbms_output.put_line('amount'||amount);
--        dbms_lob.read(my_clob, amount, prev, v_str);
--      end if;
--      dbms_output.put_line(v_str);
--      prev := indx+size_of_nl;
--      if (amount = 32767) then prev := prev-size_of_nl-1; end if;
--    end loop;
--    if (jsonp is not null) then dbms_output.put_line(')'); end if;
--  end;
--
  
  procedure htp_output_clob(my_clob clob, jsonp varchar2 default null) as
--    amount number := 4096;
--    pos number := 1;
--    len number;
--    
    l_amt    number default 4096;
    l_off   number default 1;
    l_str   varchar2(32000);
  begin
    if (jsonp is not null) then htp.prn(jsonp||'('); end if;
    
    begin
      loop
        dbms_lob.read( my_clob, l_amt, l_off, l_str );
        
        -- it is vital to use htp.PRN to avoid
        -- spurious line feeds getting added to your
        -- document
        htp.prn( l_str  );
        l_off := l_off+l_amt;
      end loop;
    exception
      when no_data_found then NULL;
    end;
    
--    
--    len := dbms_lob.getlength(my_clob);
--    
--    while (pos < len) loop
--      htp.prn(dbms_lob.substr(my_clob, amount, pos)); -- should I replace substr with dbms_lob.read 
--      --dbms_output.put_line(dbms_lob.substr(my_clob, amount, pos));
--      pos := pos + amount;
--    end loop;
--    
    if (jsonp is not null) then htp.prn(')'); end if;
  end;

end pljson_printer;
/
create or replace package body pljson_ext as
  scanner_exception exception;
  pragma exception_init(scanner_exception, -20100);
  parser_exception exception;
  pragma exception_init(parser_exception, -20101);
  jext_exception exception;
  pragma exception_init(jext_exception, -20110);

  --extra function checks if number has no fraction
  function is_integer(v pljson_value) return boolean as
    num number;
    num_double binary_double;
    int_number number(38); --the oracle way to specify an integer
    int_double binary_double; --the oracle way to specify an integer
  begin
--    
--    if (v.is_number()) then
--      myint := v.get_number();
--      return (myint = v.get_number()); --no rounding errors 
--    else
--      return false;
--    end if;
--    
    if (not v.is_number()) then
      raise_application_error(-20109, 'not a number-value');
    end if;
--     E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
    if (v.is_number_repr_number()) then
      num := v.get_number();
      int_number := trunc(num);
      --dbms_output.put_line('number: ' || num || ' -> ' || int_number);
      return (int_number = num); --no rounding errors 
    elsif (v.is_number_repr_double()) then
      num_double := v.get_double();
      int_double := trunc(num_double);
      --dbms_output.put_line('double: ' || num_double || ' -> ' || int_double);
      return (int_double = num_double); --no rounding errors 
    else
      return false;
    end if;
  end;

  --extension enables json to store dates without compromising the implementation
  function to_json_value(d date) return pljson_value as
  begin
    return pljson_value(to_char(d, format_string));
  end;

  --notice that a date type in json is also a varchar2
  function is_date(v pljson_value) return boolean as
    temp date;
  begin
    temp := pljson_ext.to_date(v);
    return true;
  exception
    when others then
      return false;
  end;

  --conversion is needed to extract dates
  function to_date(v pljson_value) return date as
  begin
    if (v.is_string()) then
      return standard.to_date(v.get_string(), format_string);
    else
      raise_application_error(-20110, 'Anydata did not contain a date-value');
    end if;
  exception
    when others then
      raise_application_error(-20110, 'Anydata did not contain a date on the format: '||format_string);
  end;

  function to_date2(v pljson_value) return date as
  begin
    return to_date(v);
  end;

--  
--    assumes single base64 string or broken into equal length lines of max 64 or 76 chars
--    (as specified by RFC-1421 or RFC-2045)
--    line ending can be CR+NL or NL
--  
  function decodeBase64Clob2Blob(p_clob clob) return blob
  is
    r_blob blob;
    clob_size number;
    pos number;
    c_buf varchar2(32767);
    r_buf raw(32767);
    v_read_size number;
    v_line_size number;
  begin
    dbms_lob.createtemporary(r_blob, false, dbms_lob.call);
--    
--      E.I.Sarmas (github.com/dsnz)   2017-12-07   fix for alignment issues
--      assumes single base64 string or broken into equal length lines of max 64 or 76 followed by CR+NL
--      as specified by RFC-1421 or RFC-2045 which seem to be the supported ones by Oracle utl_encode
--      also support single NL instead of CR+NL !
--    
    clob_size := dbms_lob.getlength(p_clob);
    v_line_size := 64;
    if clob_size >= 65 and dbms_lob.substr(p_clob, 1, 65) = chr(10) then
      v_line_size := 65;
    elsif clob_size >= 66 and dbms_lob.substr(p_clob, 1, 65) = chr(13) then
      v_line_size := 66;
    elsif clob_size >= 77 and dbms_lob.substr(p_clob, 1, 77) = chr(10) then
      v_line_size := 77;
    elsif clob_size >= 78 and dbms_lob.substr(p_clob, 1, 77) = chr(13) then
      v_line_size := 78;
    end if;
    --dbms_output.put_line('decoding in multiples of ' || v_line_size);
    v_read_size := floor(32767/v_line_size)*v_line_size;
    
    pos := 1;
    while (pos < clob_size) loop
      dbms_lob.read(p_clob, v_read_size, pos, c_buf);
      r_buf := utl_encode.base64_decode(utl_raw.cast_to_raw(c_buf));
      dbms_lob.writeappend(r_blob, utl_raw.length(r_buf), r_buf);
      pos := pos + v_read_size;
    end loop;
    return r_blob;
  end decodeBase64Clob2Blob;

--  
--    encoding in lines of 64 chars ending with CR+NL
--  
  function encodeBase64Blob2Clob(p_blob in  blob) return clob
  is
    r_clob clob;
--     E.I.Sarmas (github.com/dsnz)   2017-12-07   NOTE: must be multiple of 48 !!! 
    c_step pls_integer := 12000;
    c_buf varchar2(32767);
  begin
    if p_blob is not null then
      dbms_lob.createtemporary(r_clob, false, dbms_lob.call);
      for i in 0 .. trunc((dbms_lob.getlength(p_blob) - 1)/c_step) loop
        c_buf := utl_raw.cast_to_varchar2(utl_encode.base64_encode(dbms_lob.substr(p_blob, c_step, i * c_step + 1)));
--        
--          E.I.Sarmas (github.com/dsnz)   2017-12-07   fix for alignment issues
--          must output CR+NL at end always, so will align with the following block and can be decoded correctly
--          assumes ending in CR+NL
--        
        if substr(c_buf, length(c_buf)) != chr(10) then
          c_buf := c_buf || CHR(13) || CHR(10);
        end if;
--        
--        dbms_output.put_line(
--          'l=' || length(c_buf) ||
--          ' e=' || ascii(substr(c_buf, length(c_buf) - 1)) || ' ' || ascii(substr(c_buf, length(c_buf)))
--        );
--        
        dbms_lob.writeappend(lob_loc => r_clob, amount => length(c_buf), buffer => c_buf);
      end loop;
    end if;
    return r_clob;
  end encodeBase64Blob2Clob;

  --Json Path parser
  function parsePath(json_path varchar2, base number default 1) return pljson_list as
    build_path varchar2(32767) := '[';
    buf varchar2(4);
    endstring varchar2(1);
    indx number := 1;
    ret pljson_list;

    procedure next_char as
    begin
      if (indx <= length(json_path)) then
        buf := substr(json_path, indx, 1);
        indx := indx + 1;
      else
        buf := null;
      end if;
    end;
    --skip ws
    procedure skipws as begin while (buf in (chr(9), chr(10), chr(13), ' ')) loop next_char; end loop; end;

  begin
    next_char();
    while (buf is not null) loop
      if (buf = '.') then
        next_char();
        if (buf is null) then raise_application_error(-20110, 'JSON Path parse error: . is not a valid json_path end'); end if;
        if (not regexp_like(buf, '^[[:alnum:]\_ ]+', 'c') ) then
          raise_application_error(-20110, 'JSON Path parse error: alpha-numeric character or space expected at position '||indx);
        end if;

        if (build_path != '[') then build_path := build_path || ','; end if;
        build_path := build_path || '"';
        while (regexp_like(buf, '^[[:alnum:]\_ ]+', 'c') ) loop
          build_path := build_path || buf;
          next_char();
        end loop;
        build_path := build_path || '"';
      elsif (buf = '[') then
        next_char();
        skipws();
        if (buf is null) then raise_application_error(-20110, 'JSON Path parse error: [ is not a valid json_path end'); end if;
        if (buf in ('1','2','3','4','5','6','7','8','9') or (buf = '0' and base = 0)) then
          if (build_path != '[') then build_path := build_path || ','; end if;
          while (buf in ('0','1','2','3','4','5','6','7','8','9')) loop
            build_path := build_path || buf;
            next_char();
          end loop;
        elsif (regexp_like(buf, '^(\"|\'')', 'c')) then
          endstring := buf;
          if (build_path != '[') then build_path := build_path || ','; end if;
          build_path := build_path || '"';
          next_char();
          if (buf is null) then raise_application_error(-20110, 'JSON Path parse error: premature json_path end'); end if;
          while (buf != endstring) loop
            build_path := build_path || buf;
            next_char();
            if (buf is null) then raise_application_error(-20110, 'JSON Path parse error: premature json_path end'); end if;
            if (buf = '\') then
              next_char();
              build_path := build_path || '\' || buf;
              next_char();
            end if;
          end loop;
          build_path := build_path || '"';
          next_char();
        else
          raise_application_error(-20110, 'JSON Path parse error: expected a string or an positive integer at '||indx);
        end if;
        skipws();
        if (buf is null) then raise_application_error(-20110, 'JSON Path parse error: premature json_path end'); end if;
        if (buf != ']') then raise_application_error(-20110, 'JSON Path parse error: no array ending found. found: '|| buf); end if;
        next_char();
        skipws();
      elsif (build_path = '[') then
        if (not regexp_like(buf, '^[[:alnum:]\_ ]+', 'c') ) then
          raise_application_error(-20110, 'JSON Path parse error: alpha-numeric character or space expected at position '||indx);
        end if;
        build_path := build_path || '"';
        while (regexp_like(buf, '^[[:alnum:]\_ ]+', 'c') ) loop
          build_path := build_path || buf;
          next_char();
        end loop;
        build_path := build_path || '"';
      else
        raise_application_error(-20110, 'JSON Path parse error: expected . or [ found '|| buf || ' at position '|| indx);
      end if;

    end loop;

    build_path := build_path || ']';
    build_path := replace(replace(replace(replace(replace(build_path, chr(9), '\t'), chr(10), '\n'), chr(13), '\f'), chr(8), '\b'), chr(14), '\r');

    ret := pljson_list(build_path);
    if (base != 1) then
      --fix base 0 to base 1
      declare
        elem pljson_value;
      begin
        for i in 1 .. ret.count loop
          elem := ret.get(i);
          if (elem.is_number()) then
            ret.replace(i, elem.get_number()+1);
          end if;
        end loop;
      end;
    end if;

    return ret;
  end parsePath;

  --JSON Path getters
  function get_json_value(obj pljson, v_path varchar2, base number default 1) return pljson_value as
    path pljson_list;
    ret pljson_value;
    o pljson; l pljson_list;
  begin
    path := parsePath(v_path, base);
    ret := obj.to_json_value;
    if (path.count = 0) then return ret; end if;

    for i in 1 .. path.count loop
      if (path.get(i).is_string()) then
        --string fetch only on json
        o := pljson(ret);
        ret := o.get(path.get(i).get_string());
      else
        --number fetch on json and json_list
        if (ret.is_array()) then
          l := pljson_list(ret);
          ret := l.get(path.get(i).get_number());
        else
          o := pljson(ret);
          l := o.get_values();
          ret := l.get(path.get(i).get_number());
        end if;
      end if;
    end loop;

    return ret;
  exception
    when scanner_exception then raise;
    when parser_exception then raise;
    when jext_exception then raise;
    when others then return null;
  end get_json_value;

  --JSON Path getters
  function get_string(obj pljson, path varchar2, base number default 1) return varchar2 as
    temp pljson_value;
  begin
    temp := get_json_value(obj, path, base);
    if (temp is null or not temp.is_string()) then
      return null;
    else
      return temp.get_string();
    end if;
  end;

  function get_number(obj pljson, path varchar2, base number default 1) return number as
    temp pljson_value;
  begin
    temp := get_json_value(obj, path, base);
    if (temp is null or not temp.is_number()) then
      return null;
    else
      return temp.get_number();
    end if;
  end;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function get_double(obj pljson, path varchar2, base number default 1) return binary_double as
    temp pljson_value;
  begin
    temp := get_json_value(obj, path, base);
    if (temp is null or not temp.is_number()) then
      return null;
    else
      return temp.get_double();
    end if;
  end;

  function get_json(obj pljson, path varchar2, base number default 1) return pljson as
    temp pljson_value;
  begin
    temp := get_json_value(obj, path, base);
    if (temp is null or not temp.is_object()) then
      return null;
    else
      return pljson(temp);
    end if;
  end;

  function get_json_list(obj pljson, path varchar2, base number default 1) return pljson_list as
    temp pljson_value;
  begin
    temp := get_json_value(obj, path, base);
    if (temp is null or not temp.is_array()) then
      return null;
    else
      return pljson_list(temp);
    end if;
  end;

  function get_bool(obj pljson, path varchar2, base number default 1) return boolean as
    temp pljson_value;
  begin
    temp := get_json_value(obj, path, base);
    if (temp is null or not temp.is_bool()) then
      return null;
    else
      return temp.get_bool();
    end if;
  end;

  function get_date(obj pljson, path varchar2, base number default 1) return date as
    temp pljson_value;
  begin
    temp := get_json_value(obj, path, base);
    if (temp is null or not is_date(temp)) then
      return null;
    else
      return pljson_ext.to_date(temp);
    end if;
  end;

--   JSON Path putter internal function 
  procedure put_internal(obj in out nocopy pljson, v_path varchar2, elem pljson_value, base number) as
    val pljson_value := elem;
    path pljson_list;
    backreference pljson_list := pljson_list();

    keyval pljson_value; keynum number; keystring varchar2(4000);
    temp pljson_value := obj.to_json_value;
    obj_temp  pljson;
    list_temp pljson_list;
    inserter pljson_value;
  begin
    path := pljson_ext.parsePath(v_path, base);
    if (path.count = 0) then raise_application_error(-20110, 'PLJSON_EXT put error: cannot put with empty string.'); end if;

    --build backreference
    for i in 1 .. path.count loop
      --backreference.print(false);
      keyval := path.get(i);
      if (keyval.is_number()) then
        --number index
        keynum := keyval.get_number();
        if ((not temp.is_object()) and (not temp.is_array())) then
          if (val is null) then return; end if;
          backreference.remove_last;
          temp := pljson_list().to_json_value;
          backreference.append(temp);
        end if;

        if (temp.is_object()) then
          obj_temp := pljson(temp);
          if (obj_temp.count < keynum) then
            if (val is null) then return; end if;
            raise_application_error(-20110, 'PLJSON_EXT put error: access object with too few members.');
          end if;
          temp := obj_temp.get(keynum);
        else
          list_temp := pljson_list(temp);
          if (list_temp.count < keynum) then
            if (val is null) then return; end if;
            --raise error or quit if val is null
            for i in list_temp.count+1 .. keynum loop
              list_temp.append(pljson_value());
            end loop;
            backreference.remove_last;
            backreference.append(list_temp);
          end if;

          temp := list_temp.get(keynum);
        end if;
      else
        --string index
        keystring := keyval.get_string();
        if (not temp.is_object()) then
          --backreference.print;
          if (val is null) then return; end if;
          backreference.remove_last;
          temp := pljson().to_json_value;
          backreference.append(temp);
          --raise_application_error(-20110, 'PLJSON_EXT put error: trying to access a non object with a string.');
        end if;
        obj_temp := pljson(temp);
        temp := obj_temp.get(keystring);
      end if;

      if (temp is null) then
        if (val is null) then return; end if;
        --what to expect 
        keyval := path.get(i+1);
        if (keyval is not null and keyval.is_number()) then
          temp := pljson_list().to_json_value;
        else
          temp := pljson().to_json_value;
        end if;
      end if;
      backreference.append(temp);
    end loop;

    --  backreference.print(false);
    --  path.print(false);

    --use backreference and path together
    inserter := val;
    for i in reverse 1 .. backreference.count loop
      -- inserter.print(false);
      if ( i = 1 ) then
        keyval := path.get(1);
        if (keyval.is_string()) then
          keystring := keyval.get_string();
        else
          keynum := keyval.get_number();
          declare
            t1 pljson_value := obj.get(keynum);
          begin
            keystring := t1.mapname;
          end;
        end if;
        if (inserter is null) then obj.remove(keystring); else obj.put(keystring, inserter); end if;
      else
        temp := backreference.get(i-1);
        if (temp.is_object()) then
          keyval := path.get(i);
          obj_temp := pljson(temp);
          if (keyval.is_string()) then
            keystring := keyval.get_string();
          else
            keynum := keyval.get_number();
            declare
              t1 pljson_value := obj_temp.get(keynum);
            begin
              keystring := t1.mapname;
            end;
          end if;
          if (inserter is null) then
            obj_temp.remove(keystring);
            if (obj_temp.count > 0) then inserter := obj_temp.to_json_value; end if;
          else
            obj_temp.put(keystring, inserter);
            inserter := obj_temp.to_json_value;
          end if;
        else
          --array only number
          keynum := path.get(i).get_number();
          list_temp := pljson_list(temp);
          list_temp.remove(keynum);
          if (not inserter is null) then
            list_temp.append(inserter, keynum);
            inserter := list_temp.to_json_value;
          else
            if (list_temp.count > 0) then inserter := list_temp.to_json_value; end if;
          end if;
        end if;
      end if;

    end loop;

  end put_internal;

--   JSON Path putters 
  procedure put(obj in out nocopy pljson, path varchar2, elem varchar2, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, pljson_value(elem), base);
    end if;
  end;

  procedure put(obj in out nocopy pljson, path varchar2, elem number, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, pljson_value(elem), base);
    end if;
  end;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure put(obj in out nocopy pljson, path varchar2, elem binary_double, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, pljson_value(elem), base);
    end if;
  end;

  procedure put(obj in out nocopy pljson, path varchar2, elem pljson, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, elem.to_json_value, base);
    end if;
  end;

  procedure put(obj in out nocopy pljson, path varchar2, elem pljson_list, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, elem.to_json_value, base);
    end if;
  end;

  procedure put(obj in out nocopy pljson, path varchar2, elem boolean, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, pljson_value(elem), base);
    end if;
  end;

  procedure put(obj in out nocopy pljson, path varchar2, elem pljson_value, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, elem, base);
    end if;
  end;

  procedure put(obj in out nocopy pljson, path varchar2, elem date, base number default 1) as
  begin
    if elem is null then
      put_internal(obj, path, pljson_value(), base);
    else
      put_internal(obj, path, pljson_ext.to_json_value(elem), base);
    end if;
  end;

  procedure remove(obj in out nocopy pljson, path varchar2, base number default 1) as
  begin
    pljson_ext.put_internal(obj, path, null, base);
--    if (json_ext.get_json_value(obj, path) is not null) then
--    end if;
  end remove;

  --Pretty print with JSON Path
  function pp(obj pljson, v_path varchar2) return varchar2 as
    json_part pljson_value;
  begin
    json_part := pljson_ext.get_json_value(obj, v_path);
    if (json_part is null) then
      return '';
    else
      return pljson_printer.pretty_print_any(json_part); --escapes a possible internal string
    end if;
  end pp;

  procedure pp(obj pljson, v_path varchar2) as --using dbms_output.put_line
  begin
    dbms_output.put_line(pp(obj, v_path));
  end pp;

  -- spaces = false!
  procedure pp_htp(obj pljson, v_path varchar2) as --using htp.print
    json_part pljson_value;
  begin
    json_part := pljson_ext.get_json_value(obj, v_path);
    if (json_part is null) then
      htp.print;
    else
      htp.print(pljson_printer.pretty_print_any(json_part, false));
    end if;
  end pp_htp;

  function base64(binarydata blob) return pljson_list as
    obj pljson_list := pljson_list();
    c clob := empty_clob();

    v_clob_offset NUMBER := 1;
    v_lang_context NUMBER := DBMS_LOB.DEFAULT_LANG_CTX;
    v_amount PLS_INTEGER;
  begin
    dbms_lob.createtemporary(c, false, dbms_lob.call);
    c := encodeBase64Blob2Clob(binarydata);
    v_amount := DBMS_LOB.GETLENGTH(c);
    v_clob_offset := 1;
    --dbms_output.put_line('V amount: '||v_amount);
    while (v_clob_offset < v_amount) loop
      --dbms_output.put_line(v_offset);
      --temp := ;
      --dbms_output.put_line('size: '||length(temp));
      obj.append(dbms_lob.SUBSTR(c, 4000, v_clob_offset));
      v_clob_offset := v_clob_offset + 4000;
    end loop;
    dbms_lob.freetemporary(c);
  --dbms_output.put_line(obj.count);
  --dbms_output.put_line(obj.get_last().to_char);
    return obj;

  end base64;

  function base64(l pljson_list) return blob as
    c clob := empty_clob();
    b_ret blob;

    v_lang_context NUMBER := 0; --DBMS_LOB.DEFAULT_LANG_CTX;
--    v_amount PLS_INTEGER;
  begin
    dbms_lob.createtemporary(c, false, dbms_lob.call);
    for i in 1 .. l.count loop
      dbms_lob.append(c, l.get(i).get_string());
    end loop;
--    v_amount := DBMS_LOB.GETLENGTH(c);
--    dbms_output.put_line('L C'||v_amount);
    b_ret := decodeBase64Clob2Blob(c);
    dbms_lob.freetemporary(c);
    return b_ret;
  end base64;

  function encode(binarydata blob) return pljson_value as
    obj pljson_value;
    c clob;
    v_lang_context NUMBER := DBMS_LOB.DEFAULT_LANG_CTX;
  begin
    dbms_lob.createtemporary(c, false, dbms_lob.call);
    c := encodeBase64Blob2Clob(binarydata);
    obj := pljson_value(c);

  --dbms_output.put_line(obj.count);
  --dbms_output.put_line(obj.get_last().to_char);
--    dbms_lob.freetemporary(c);
    return obj;
  end encode;

  function decode(v pljson_value) return blob as
    --c clob := empty_clob();
    c clob;
    b_ret blob;

    v_lang_context NUMBER := 0; --DBMS_LOB.DEFAULT_LANG_CTX;
--    v_amount PLS_INTEGER;
  begin
--    
--    dbms_lob.createtemporary(c, false, dbms_lob.call);
--    v.get_string(c);
--    
    c := v.get_clob();
--    v_amount := DBMS_LOB.GETLENGTH(c);
--    dbms_output.put_line('L C'||v_amount);
    b_ret := decodeBase64Clob2Blob(c);
--    dbms_lob.freetemporary(c);
    return b_ret;

  end decode;

  procedure blob2clob(b blob, c out clob, charset varchar2 default 'UTF8') as
    v_dest_offset integer := 1;
    v_src_offset integer := 1;
    v_lang_context integer := 0;
    v_warning integer := 0;
  begin
    dbms_lob.createtemporary(c, false, dbms_lob.call);
    dbms_lob.converttoclob(
      dest_lob => c,
      src_blob => b,
      amount => dbms_lob.LOBMAXSIZE,
      dest_offset => v_dest_offset,
      src_offset => v_src_offset,
      blob_csid => nls_charset_id(charset),
      lang_context => v_lang_context,
      warning => v_warning);
  end;
end pljson_ext;
/
create or replace type body pljson_value as

  constructor function pljson_value(elem pljson_element) return self as result as
  begin
    case
      when elem is of (pljson)      then self.typeval := 1;
      when elem is of (pljson_list) then self.typeval := 2;
      else raise_application_error(-20102, 'PLJSON_VALUE init error (PLJSON or PLJSON_LIST allowed)');
    end case;
    self.object_or_array := elem;
    if(self.object_or_array is null) then self.typeval := 6; end if;

    return;
  end pljson_value;

  constructor function pljson_value(str varchar2, esc boolean default true) return self as result as
  begin
    self.typeval := 3;
    if(esc) then self.num := 1; else self.num := 0; end if; --message to pretty printer
    self.str := str;
    return;
  end pljson_value;

  constructor function pljson_value(str clob, esc boolean default true) return self as result as
--     E.I.Sarmas (github.com/dsnz)   2016-01-21   limit to 5000 chars 
    amount number := 32767; -- for Unicode text, varchar2 'self.str' not exceed 5000 chars, does not limit size of data 
    --by:zyk,目前没有使用unicode，从5000改为32767
  begin
    self.typeval := 3;
    if(esc) then self.num := 1; else self.num := 0; end if; --message to pretty printer
    if(dbms_lob.getlength(str) > amount) then
      extended_str := str;
    end if;
    -- GHS 20120615: Added IF structure to handle null clobs
    if dbms_lob.getlength(str) > 0 then
      dbms_lob.read(str, amount, 1, self.str);
    end if;
    return;
  end pljson_value;

  constructor function pljson_value(num number) return self as result as
  begin
    self.typeval := 4;
    self.num := num;
--     E.I.Sarmas (github.com/dsnz)   2016-11-03   support for binary_double numbers; typeval not changed, it is still json number 
    self.num_repr_number_p := 't';
    self.num_double := num;
    if (to_number(self.num_double) = self.num) then
      self.num_repr_double_p := 't';
    else
      self.num_repr_double_p := 'f';
    end if;
--     
    if(self.num is null) then self.typeval := 6; end if;
    return;
  end pljson_value;

--   E.I.Sarmas (github.com/dsnz)   2016-11-03   support for binary_double numbers; typeval not changed, it is still json number 
  constructor function pljson_value(num_double binary_double) return self as result as
  begin
    self.typeval := 4;
    self.num_double := num_double;
    self.num_repr_double_p := 't';
    self.num := num_double;
    if (to_binary_double(self.num) = self.num_double) then
      self.num_repr_number_p := 't';
    else
      self.num_repr_number_p := 'f';
    end if;
    if(self.num_double is null) then self.typeval := 6; end if;
    return;
  end pljson_value;

  constructor function pljson_value(b boolean) return self as result as
  begin
    self.typeval := 5;
    self.num := 0;
    if(b) then self.num := 1; end if;
    if(b is null) then self.typeval := 6; end if;
    return;
  end pljson_value;

  constructor function pljson_value return self as result as
  begin
    self.typeval := 6; -- for JSON null 
    return;
  end pljson_value;

  member function get_element return pljson_element as
  begin
    if (self.typeval in (1,2)) then
      return self.object_or_array;
    end if;
    return null;
  end get_element;

  static function makenull return pljson_value as
  begin
    return pljson_value;
  end makenull;

  member function get_type return varchar2 as
  begin
    case self.typeval
    when 1 then return 'object';
    when 2 then return 'array';
    when 3 then return 'string';
    when 4 then return 'number';
    when 5 then return 'bool';
    when 6 then return 'null';
    end case;

    return 'unknown type';
  end get_type;

  member function get_string(max_byte_size number default null, max_char_size number default null) return varchar2 as
  begin
    if(self.typeval = 3) then
      if(max_byte_size is not null) then
        return substrb(self.str,1,max_byte_size);
      elsif (max_char_size is not null) then
        return substr(self.str,1,max_char_size);
      else
        return self.str;
      end if;
    end if;
    return null;
  end get_string;

  member procedure get_string(self in pljson_value, buf in out nocopy clob) as
  begin
    if(self.typeval = 3) then
      if(extended_str is not null) then
        dbms_lob.copy(buf, extended_str, dbms_lob.getlength(extended_str));
      else
        dbms_lob.writeappend(buf, length(self.str), self.str);
      end if;
    end if;
  end get_string;

  member function get_clob return clob as
  begin
    if(self.typeval = 3) then
      if(extended_str is not null) then
        --dbms_lob.copy(buf, extended_str, dbms_lob.getlength(extended_str));
        return self.extended_str;
      else
        --dbms_lob.writeappend(buf, length(self.str), self.str);
        return self.str;
      end if;
    end if;
  end get_clob;

  member function get_number return number as
  begin
    if(self.typeval = 4) then
      return self.num;
    end if;
    return null;
  end get_number;

--   E.I.Sarmas (github.com/dsnz)   2016-11-03   support for binary_double numbers 
  member function get_double return binary_double as
  begin
    if(self.typeval = 4) then
      return self.num_double;
    end if;
    return null;
  end get_double;

  member function get_bool return boolean as
  begin
    if(self.typeval = 5) then
      return self.num = 1;
    end if;
    return null;
  end get_bool;

  member function get_null return varchar2 as
  begin
    if(self.typeval = 6) then
      return 'null';
    end if;
    return null;
  end get_null;

  member function is_object return boolean as begin return self.typeval = 1; end;
  member function is_array return boolean as begin return self.typeval = 2; end;
  member function is_string return boolean as begin return self.typeval = 3; end;
  member function is_number return boolean as begin return self.typeval = 4; end;
  member function is_bool return boolean as begin return self.typeval = 5; end;
  member function is_null return boolean as begin return self.typeval = 6; end;

--   E.I.Sarmas (github.com/dsnz)   2016-11-03   support for binary_double numbers, is_number is still true, extra check 
--   return true if 'number' is representable by Oracle number 
  member function is_number_repr_number return boolean is
  begin
    if self.typeval != 4 then
      return false;
    end if;
    return (num_repr_number_p = 't');
  end;

--   return true if 'number' is representable by Oracle binary_double 
  member function is_number_repr_double return boolean is
  begin
    if self.typeval != 4 then
      return false;
    end if;
    return (num_repr_double_p = 't');
  end;

--   E.I.Sarmas (github.com/dsnz)   2016-11-03   support for binary_double numbers 
  -- set value for number from string representation; to replace to_number in pljson_parser
  -- can automatically decide and use binary_double if needed (set repr variables)
  -- underflows and overflows count as representable if happen on both type representations
  -- less confusing than new constructor with dummy argument for overloading
  -- centralized parse_number to use everywhere else and replace code in pljson_parser
  --
  -- WARNING:
  --
  -- procedure does not work correctly if called standalone in locales that
  -- use a character other than "." for decimal point
  --
  -- parse_number() is intended to be used inside pljson_parser which
  -- uses session NLS_PARAMETERS to get decimal point and
  -- changes "." to this decimal point before calling parse_number()
  --
  member procedure parse_number(str varchar2) is
  begin
    if self.typeval != 4 then
      return;
    end if;
    self.num := to_number(str);
    self.num_repr_number_p := 't';
    self.num_double := to_binary_double(str);
    self.num_repr_double_p := 't';
    if (to_binary_double(self.num) != self.num_double) then
      self.num_repr_number_p := 'f';
    end if;
    if (to_number(self.num_double) != self.num) then
      self.num_repr_double_p := 'f';
    end if;
  end parse_number;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  -- centralized toString to use everywhere else and replace code in pljson_printer
  member function number_toString return varchar2 is
    num number;
    num_double binary_double;
    buf varchar2(4000);
  begin
--     unrolled, instead of using two nested fuctions for speed 
    if (self.num_repr_number_p = 't') then
      num := self.num;
      if (num > 1e127d) then
        return '1e309'; -- json representation of infinity ! 
      end if;
      if (num < -1e127d) then
        return '-1e309'; -- json representation of infinity ! 
      end if;
      buf := STANDARD.to_char(num, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
      if (-1 < num and num < 0 and substr(buf, 1, 2) = '-.') then
        buf := '-0' || substr(buf, 2);
      elsif (0 < num and num < 1 and substr(buf, 1, 1) = '.') then
        buf := '0' || buf;
      end if;
      return buf;
    else
      num_double := self.num_double;
      if (num_double = +BINARY_DOUBLE_INFINITY) then
        return '1e309'; -- json representation of infinity ! 
      end if;
      if (num_double = -BINARY_DOUBLE_INFINITY) then
        return '-1e309'; -- json representation of infinity ! 
      end if;
      buf := STANDARD.to_char(num_double, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''');
      if (-1 < num_double and num_double < 0 and substr(buf, 1, 2) = '-.') then
        buf := '-0' || substr(buf, 2);
      elsif (0 < num_double and num_double < 1 and substr(buf, 1, 1) = '.') then
        buf := '0' || buf;
      end if;
      return buf;
    end if;
  end number_toString;

--   Output methods 
  member function to_char(spaces boolean default true, chars_per_line number default 0) return varchar2 as
  begin
    if(spaces is null) then
      return pljson_printer.pretty_print_any(self, line_length => chars_per_line);
    else
      return pljson_printer.pretty_print_any(self, spaces, line_length => chars_per_line);
    end if;
  end;

  member procedure to_clob(self in pljson_value, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true) as
  begin
    if(spaces is null) then
      pljson_printer.pretty_print_any(self, false, buf, line_length => chars_per_line, erase_clob => erase_clob);
    else
      pljson_printer.pretty_print_any(self, spaces, buf, line_length => chars_per_line, erase_clob => erase_clob);
    end if;
  end;

  member procedure print(self in pljson_value, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null) as --32512 is the real maximum in sqldeveloper
    my_clob clob;
  begin
    my_clob := empty_clob();
    dbms_lob.createtemporary(my_clob, true);
    pljson_printer.pretty_print_any(self, spaces, my_clob, case when (chars_per_line>32512) then 32512 else chars_per_line end);
    pljson_printer.dbms_output_clob(my_clob, pljson_printer.newline_char, jsonp);
    dbms_lob.freetemporary(my_clob);
  end;

  member procedure htp(self in pljson_value, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null) as
    my_clob clob;
  begin
    my_clob := empty_clob();
    dbms_lob.createtemporary(my_clob, true);
    pljson_printer.pretty_print_any(self, spaces, my_clob, chars_per_line);
    pljson_printer.htp_output_clob(my_clob, jsonp);
    dbms_lob.freetemporary(my_clob);
  end;

  member function value_of(self in pljson_value, max_byte_size number default null, max_char_size number default null) return varchar2 as
  begin
    case self.typeval
    when 1 then return 'json object';
    when 2 then return 'json array';
    when 3 then return self.get_string(max_byte_size, max_char_size);
    when 4 then return self.get_number();
    when 5 then if(self.get_bool()) then return 'true'; else return 'false'; end if;
    else return null;
    end case;
  end;

end;
/
create or replace type body pljson_list as

--   constructors 
  constructor function pljson_list return self as result as
  begin
    self.list_data := pljson_value_array();
    return;
  end;

  constructor function pljson_list(str varchar2) return self as result as
  begin
    self := pljson_parser.parse_list(str);
    return;
  end;

  constructor function pljson_list(str clob) return self as result as
  begin
    self := pljson_parser.parse_list(str);
    return;
  end;

  constructor function pljson_list(str blob, charset varchar2 default 'UTF8') return self as result as
    c_str clob;
  begin
    pljson_ext.blob2clob(str, c_str, charset);
    self := pljson_parser.parse_list(c_str);
    dbms_lob.freetemporary(c_str);
    return;
  end;
  --lshuo注释
  --由于10g不支持DESC_TAB3 ，且无法返回pljson_narray，pljson_varray
	constructor function pljson_list(str_array pljson_varray) return self as result as
	begin
	  self.list_data := pljson_value_array();
	  for i in str_array.FIRST .. str_array.LAST loop
	    append(str_array(i));
	  end loop;
	  return;
	end;
	constructor function pljson_list(num_array pljson_narray) return self as result as
	begin
	  self.list_data := pljson_value_array();
	  for i in num_array.FIRST .. num_array.LAST loop
	    append(num_array(i));
	  end loop;
	  return;
	end;	
  constructor function pljson_list(elem pljson_value) return self as result as
  begin
    self := treat(elem.object_or_array as pljson_list);
    return;
  end;

--   list management 
  member procedure append(self in out nocopy pljson_list, elem pljson_value, position pls_integer default null) as
    indx pls_integer;
    insert_value pljson_value;
  begin
    insert_value := elem;
    if insert_value is null then
      insert_value := pljson_value();
    end if;
    if (position is null or position > self.count) then --end of list
      indx := self.count + 1;
      self.list_data.extend(1);
      self.list_data(indx) := insert_value;
    elsif (position < 1) then --new first
      indx := self.count;
      self.list_data.extend(1);
      for x in reverse 1 .. indx loop
        self.list_data(x+1) := self.list_data(x);
      end loop;
      self.list_data(1) := insert_value;
    else
      indx := self.count;
      self.list_data.extend(1);
      for x in reverse position .. indx loop
        self.list_data(x+1) := self.list_data(x);
      end loop;
      self.list_data(position) := insert_value;
    end if;
  end;

  member procedure append(self in out nocopy pljson_list, elem varchar2, position pls_integer default null) as
  begin
    append(pljson_value(elem), position);
  end;

  member procedure append(self in out nocopy pljson_list, elem clob, position pls_integer default null) as
  begin
    append(pljson_value(elem), position);
  end;

  member procedure append(self in out nocopy pljson_list, elem number, position pls_integer default null) as
  begin
    if (elem is null) then
      append(pljson_value(), position);
    else
      append(pljson_value(elem), position);
    end if;
  end;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure append(self in out nocopy pljson_list, elem binary_double, position pls_integer default null) as
  begin
    if (elem is null) then
      append(pljson_value(), position);
    else
      append(pljson_value(elem), position);
    end if;
  end;

  member procedure append(self in out nocopy pljson_list, elem boolean, position pls_integer default null) as
  begin
    if (elem is null) then
      append(pljson_value(), position);
    else
      append(pljson_value(elem), position);
    end if;
  end;

  member procedure append(self in out nocopy pljson_list, elem pljson_list, position pls_integer default null) as
  begin
    if (elem is null) then
      append(pljson_value(), position);
    else
      append(elem.to_json_value, position);
    end if;
  end;

  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem pljson_value) as
    insert_value pljson_value;
    indx number;
  begin
    insert_value := elem;
    if insert_value is null then
      insert_value := pljson_value();
    end if;
    if (position > self.count) then --end of list
      indx := self.count + 1;
      self.list_data.extend(1);
      self.list_data(indx) := insert_value;
    elsif (position < 1) then --maybe an error message here
      null;
    else
      self.list_data(position) := insert_value;
    end if;
  end;

  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem varchar2) as
  begin
    replace(position, pljson_value(elem));
  end;

  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem clob) as
  begin
    replace(position, pljson_value(elem));
  end;

  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem number) as
  begin
    if (elem is null) then
      replace(position, pljson_value());
    else
      replace(position, pljson_value(elem));
    end if;
  end;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem binary_double) as
  begin
    if (elem is null) then
      replace(position, pljson_value());
    else
      replace(position, pljson_value(elem));
    end if;
  end;

  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem boolean) as
  begin
    if (elem is null) then
      replace(position, pljson_value());
    else
      replace(position, pljson_value(elem));
    end if;
  end;

  member procedure replace(self in out nocopy pljson_list, position pls_integer, elem pljson_list) as
  begin
    if (elem is null) then
      replace(position, pljson_value());
    else
      replace(position, elem.to_json_value);
    end if;
  end;

  member procedure remove(self in out nocopy pljson_list, position pls_integer) as
  begin
    if (position is null or position < 1 or position > self.count) then return; end if;
    for x in (position+1) .. self.count loop
      self.list_data(x-1) := self.list_data(x);
    end loop;
    self.list_data.trim(1);
  end;

  member procedure remove_first(self in out nocopy pljson_list) as
  begin
    for x in 2 .. self.count loop
      self.list_data(x-1) := self.list_data(x);
    end loop;
    if (self.count > 0) then
      self.list_data.trim(1);
    end if;
  end;

  member procedure remove_last(self in out nocopy pljson_list) as
  begin
    if (self.count > 0) then
      self.list_data.trim(1);
    end if;
  end;

  member function count return number as
  begin
    return self.list_data.count;
  end;

  member function get(position pls_integer) return pljson_value as
  begin
    if (self.count >= position and position > 0) then
      return self.list_data(position);
    end if;
    return null; 
  end;

  member function get_string(position pls_integer) return varchar2 as
    elem pljson_value := get(position);
  begin
    
    if elem is not null then
      return elem.get_string();
    end if;
    return null;
    
--    return elem.get_string();
  end;

  member function get_clob(position pls_integer) return clob as
    elem pljson_value := get(position);
  begin
    
    if elem is not null then
      return elem.get_clob();
    end if;
    return null;
    
--    return elem.get_clob();
  end;

  member function get_number(position pls_integer) return number as
    elem pljson_value := get(position);
  begin
    
    if elem is not null then
      return elem.get_number();
    end if;
    return null;
    
--    return elem.get_number();
  end;

  member function get_double(position pls_integer) return binary_double as
    elem pljson_value := get(position);
  begin
    
    if elem is not null then
      return elem.get_double();
    end if;
    return null;
    
--    return elem.get_double();
  end;

  member function get_bool(position pls_integer) return boolean as
    elem pljson_value := get(position);
  begin
    
    if elem is not null then
      return elem.get_bool();
    end if;
    return null;
    
--    return elem.get_bool();
  end;

  member function get_pljson_list(position pls_integer) return pljson_list as
    elem pljson_value := get(position);
  begin
    
    if elem is not null then
      return treat(elem.object_or_array as pljson_list);
    end if;
    return null;
    
--    return treat(elem.object_or_array as pljson_list);
  end;

  member function head return pljson_value as
  begin
    if (self.count > 0) then
      return self.list_data(self.list_data.first);
    end if;
    return null; -- do not throw error, just return null
  end;

  member function last return pljson_value as
  begin
    if (self.count > 0) then
      return self.list_data(self.list_data.last);
    end if;
    return null; -- do not throw error, just return null
  end;

  member function tail return pljson_list as
    t pljson_list;
  begin
    if (self.count > 0) then
      t := self; --pljson_list(self.to_json_value);
      t.remove(1);
      return t;
    else
      return pljson_list();
    end if;
  end;

  member function to_json_value return pljson_value as
  begin
    return pljson_value(self);
  end;

  member function to_char(spaces boolean default true, chars_per_line number default 0) return varchar2 as
  begin
    if (spaces is null) then
      return pljson_printer.pretty_print_list(self, line_length => chars_per_line);
    else
      return pljson_printer.pretty_print_list(self, spaces, line_length => chars_per_line);
    end if;
  end;

  member procedure to_clob(self in pljson_list, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true) as
  begin
    if (spaces is null) then
      pljson_printer.pretty_print_list(self, false, buf, line_length => chars_per_line, erase_clob => erase_clob);
    else
      pljson_printer.pretty_print_list(self, spaces, buf, line_length => chars_per_line, erase_clob => erase_clob);
    end if;
  end;

  member procedure print(self in pljson_list, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null) as --32512 is the real maximum in sqldeveloper
    my_clob clob;
  begin
    my_clob := empty_clob();
    dbms_lob.createtemporary(my_clob, true);
    pljson_printer.pretty_print_list(self, spaces, my_clob, case when (chars_per_line>32512) then 32512 else chars_per_line end);
    pljson_printer.dbms_output_clob(my_clob, pljson_printer.newline_char, jsonp);
    dbms_lob.freetemporary(my_clob);
  end;

  member procedure htp(self in pljson_list, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null) as
    my_clob clob;
  begin
    my_clob := empty_clob();
    dbms_lob.createtemporary(my_clob, true);
    pljson_printer.pretty_print_list(self, spaces, my_clob, chars_per_line);
    pljson_printer.htp_output_clob(my_clob, jsonp);
    dbms_lob.freetemporary(my_clob);
  end;

--   json path 
  member function path(json_path varchar2, base number default 1) return pljson_value as
    cp pljson_list := self;
  begin
    return pljson_ext.get_json_value(pljson(cp), json_path, base);
  end path;


--   json path_put 
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem pljson_value, base number default 1) as
    objlist pljson;
    jp pljson_list := pljson_ext.parsePath(json_path, base);
  begin
    while (jp.head().get_number() > self.count) loop
      self.append(pljson_value());
    end loop;

    objlist := pljson(self);
    pljson_ext.put(objlist, json_path, elem, base);
    self := objlist.get_values;
  end path_put;

  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem varchar2, base number default 1) as
    objlist pljson;
    jp pljson_list := pljson_ext.parsePath(json_path, base);
  begin
    while (jp.head().get_number() > self.count) loop
      self.append(pljson_value());
    end loop;

    objlist := pljson(self);
    pljson_ext.put(objlist, json_path, elem, base);
    self := objlist.get_values;
  end path_put;

  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem clob, base number default 1) as
    objlist pljson;
    jp pljson_list := pljson_ext.parsePath(json_path, base);
  begin
    while (jp.head().get_number() > self.count) loop
      self.append(pljson_value());
    end loop;

    objlist := pljson(self);
    pljson_ext.put(objlist, json_path, elem, base);
    self := objlist.get_values;
  end path_put;

  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem number, base number default 1) as
    objlist pljson;
    jp pljson_list := pljson_ext.parsePath(json_path, base);
  begin
    while (jp.head().get_number() > self.count) loop
      self.append(pljson_value());
    end loop;

    objlist := pljson(self);
    if (elem is null) then
      pljson_ext.put(objlist, json_path, pljson_value(), base);
    else
      pljson_ext.put(objlist, json_path, elem, base);
    end if;
    self := objlist.get_values;
  end path_put;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem binary_double, base number default 1) as
    objlist pljson;
    jp pljson_list := pljson_ext.parsePath(json_path, base);
  begin
    while (jp.head().get_number() > self.count) loop
      self.append(pljson_value());
    end loop;

    objlist := pljson(self);
    if (elem is null) then
      pljson_ext.put(objlist, json_path, pljson_value(), base);
    else
      pljson_ext.put(objlist, json_path, elem, base);
    end if;
    self := objlist.get_values;
  end path_put;

  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem boolean, base number default 1) as
    objlist pljson;
    jp pljson_list := pljson_ext.parsePath(json_path, base);
  begin
    while (jp.head().get_number() > self.count) loop
      self.append(pljson_value());
    end loop;

    objlist := pljson(self);
    if (elem is null) then
      pljson_ext.put(objlist, json_path, pljson_value(), base);
    else
      pljson_ext.put(objlist, json_path, elem, base);
    end if;
    self := objlist.get_values;
  end path_put;

  member procedure path_put(self in out nocopy pljson_list, json_path varchar2, elem pljson_list, base number default 1) as
    objlist pljson;
    jp pljson_list := pljson_ext.parsePath(json_path, base);
  begin
    while (jp.head().get_number() > self.count) loop
      self.append(pljson_value());
    end loop;

    objlist := pljson(self);
    if (elem is null) then
      pljson_ext.put(objlist, json_path, pljson_value(), base);
    else
      pljson_ext.put(objlist, json_path, elem, base);
    end if;
    self := objlist.get_values;
  end path_put;

--   json path_remove 
  member procedure path_remove(self in out nocopy pljson_list, json_path varchar2, base number default 1) as
    objlist pljson := pljson(self);
  begin
    pljson_ext.remove(objlist, json_path, base);
    self := objlist.get_values;
  end path_remove;

--   --backwards compatibility
--  member procedure add_elem(self in out nocopy json_list, elem json_value, position pls_integer default null) as begin append(elem,position); end;
--  member procedure add_elem(self in out nocopy json_list, elem varchar2, position pls_integer default null) as begin append(elem,position); end;
--  member procedure add_elem(self in out nocopy json_list, elem number, position pls_integer default null) as begin append(elem,position); end;
--  member procedure add_elem(self in out nocopy json_list, elem boolean, position pls_integer default null) as begin append(elem,position); end;
--  member procedure add_elem(self in out nocopy json_list, elem json_list, position pls_integer default null) as begin append(elem,position); end;
--
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem json_value) as begin replace(position,elem); end;
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem varchar2) as begin replace(position,elem); end;
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem number) as begin replace(position,elem); end;
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem boolean) as begin replace(position,elem); end;
--  member procedure set_elem(self in out nocopy json_list, position pls_integer, elem json_list) as begin replace(position,elem); end;
--
--  member procedure remove_elem(self in out nocopy json_list, position pls_integer) as begin remove(position); end;
--  member function get_elem(position pls_integer) return json_value as begin return get(position); end;
--  member function get_first return json_value as begin return head(); end;
--  member function get_last return json_value as begin return last(); end;
----

end;
/


create or replace type body pljson as

--   constructors 
  constructor function pljson return self as result as
  begin
    self.json_data := pljson_value_array();
    self.check_for_duplicate := 1;
    return;
  end;

  constructor function pljson(str varchar2) return self as result as
  begin
    self := pljson_parser.parser(str);
    self.check_for_duplicate := 1;
    return;
  end;

  constructor function pljson(str in clob) return self as result as
  begin
    self := pljson_parser.parser(str);
    self.check_for_duplicate := 1;
    return;
  end;

  constructor function pljson(str in blob, charset varchar2 default 'UTF8') return self as result as
    c_str clob;
  begin
    pljson_ext.blob2clob(str, c_str, charset);
    self := pljson_parser.parser(c_str);
    self.check_for_duplicate := 1;
    dbms_lob.freetemporary(c_str);
    return;
  end;

  constructor function pljson(str_array pljson_varray) return self as result as
    new_pair boolean := True;
    pair_name varchar2(32767);
    pair_value varchar2(32767);
  begin
    self.json_data := pljson_value_array();
    self.check_for_duplicate := 1;
    for i in str_array.FIRST .. str_array.LAST loop
      if new_pair then
        pair_name := str_array(i);
        new_pair := False;
      else
        pair_value := str_array(i);
        put(pair_name, pair_value);
        new_pair := True;
      end if;
    end loop;
    return;
  end;

  constructor function pljson(elem pljson_value) return self as result as
  begin
    self := treat(elem.object_or_array as pljson);
    self.check_for_duplicate := 1;
    return;
  end;

  constructor function pljson(l in out nocopy pljson_list) return self as result as
  begin
    for i in 1 .. l.list_data.count loop
      if(l.list_data(i).mapname is null or l.list_data(i).mapname like 'row%') then
      l.list_data(i).mapname := 'row'||i;
      end if;
      l.list_data(i).mapindx := i;
    end loop;

    self.json_data := l.list_data;
    self.check_for_duplicate := 1;
    return;
  end;

--   member management 
  member procedure remove(self in out nocopy pljson, pair_name varchar2) as
    temp pljson_value;
    indx pls_integer;
  begin
    temp := get(pair_name);
    if (temp is null) then return; end if;

    indx := json_data.next(temp.mapindx);
    loop
      exit when indx is null;
      json_data(indx).mapindx := indx - 1;
      json_data(indx-1) := json_data(indx);
      indx := json_data.next(indx);
    end loop;
    json_data.trim(1);
  end;

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value pljson_value, position pls_integer default null) as
    insert_value pljson_value;
    indx pls_integer; x number;
    temp pljson_value;
  begin
    --dbms_output.put_line('name: ' || pair_name);

    --if (pair_name is null) then
    --  raise_application_error(-20102, 'JSON put-method type error: name cannot be null');
    --end if;
    insert_value := pair_value;
    if insert_value is null then
      insert_value := pljson_value();
    end if;
    insert_value.mapname := pair_name;
    if (self.check_for_duplicate = 1) then temp := get(pair_name); else temp := null; end if;
    if (temp is not null) then
      insert_value.mapindx := temp.mapindx;
      json_data(temp.mapindx) := insert_value;
      return;
    elsif (position is null or position > self.count) then
      --insert at the end of the list
      --dbms_output.put_line('insert end');
      json_data.extend(1);
--       changed to common style of updating mapindx; fix bug in assignment order 
      insert_value.mapindx := json_data.count;
      json_data(json_data.count) := insert_value;
      --dbms_output.put_line('mapindx: ' || insert_value.mapindx);
      --dbms_output.put_line('mapname: ' || insert_value.mapname);
    elsif (position < 2) then
      --insert at the start of the list
      indx := json_data.last;
      json_data.extend;
      loop
        exit when indx is null;
        temp := json_data(indx);
        temp.mapindx := indx+1;
        json_data(temp.mapindx) := temp;
        indx := json_data.prior(indx);
      end loop;
--       changed to common style of updating mapindx; fix bug in assignment order 
      insert_value.mapindx := 1;
      json_data(1) := insert_value;
    else
      --insert somewhere in the list
      indx := json_data.last;
      --dbms_output.put_line('indx: ' || indx);
      json_data.extend;
      loop
        --dbms_output.put_line('indx: ' || indx);
        temp := json_data(indx);
        temp.mapindx := indx + 1;
        json_data(temp.mapindx) := temp;
        exit when indx = position;
        indx := json_data.prior(indx);
      end loop;
--       changed to common style of updating mapindx; fix bug in assignment order 
      insert_value.mapindx := position;
      json_data(position) := insert_value;
    end if;
  end;

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value varchar2, position pls_integer default null) as
  begin
    put(pair_name, pljson_value(pair_value), position);
  end;

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value clob, position pls_integer default null) as
  begin
    put(pair_name, pljson_value(pair_value), position);
  end;

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value number, position pls_integer default null) as
  begin
    if (pair_value is null) then
      put(pair_name, pljson_value(), position);
    else
      put(pair_name, pljson_value(pair_value), position);
    end if;
  end;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value binary_double, position pls_integer default null) as
  begin
    if (pair_value is null) then
      put(pair_name, pljson_value(), position);
    else
      put(pair_name, pljson_value(pair_value), position);
    end if;
  end;

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value boolean, position pls_integer default null) as
  begin
    if (pair_value is null) then
      put(pair_name, pljson_value(), position);
    else
      put(pair_name, pljson_value(pair_value), position);
    end if;
  end;

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value pljson, position pls_integer default null) as
  begin
    if (pair_value is null) then
      put(pair_name, pljson_value(), position);
    else
      put(pair_name, pair_value.to_json_value, position);
    end if;
  end;

  member procedure put(self in out nocopy pljson, pair_name varchar2, pair_value pljson_list, position pls_integer default null) as
  begin
    if (pair_value is null) then
      put(pair_name, pljson_value(), position);
    else
      put(pair_name, pair_value.to_json_value, position);
    end if;
  end;

  member function count return number as
  begin
    return self.json_data.count;
  end;

  member function get(pair_name varchar2) return pljson_value as
    indx pls_integer;
  begin
    indx := json_data.first;
    loop
      exit when indx is null;
      if (pair_name is null and json_data(indx).mapname is null) then return json_data(indx); end if;
      if (json_data(indx).mapname = pair_name) then return json_data(indx); end if;
      indx := json_data.next(indx);
    end loop;
    return null;
  end;

  member function get_string(pair_name varchar2) return varchar2 as
    elem pljson_value := get(pair_name);
  begin
    
    if elem is not null then
      return elem.get_string();
    end if;
    return null;
    
--    return elem.get_string();
  end;

  member function get_clob(pair_name varchar2) return clob as
    elem pljson_value := get(pair_name);
  begin
    
    if elem is not null then
      return elem.get_clob();
    end if;
    return null;
    
--    return elem.get_clob();
  end;

  member function get_number(pair_name varchar2) return number as
    elem pljson_value := get(pair_name);
  begin
    
    if elem is not null then
      return elem.get_number();
    end if;
    return null;
    
--    return elem.get_number();
  end;

  member function get_double(pair_name varchar2) return binary_double as
    elem pljson_value := get(pair_name);
  begin
    
    if elem is not null then
      return elem.get_double();
    end if;
    return null;
    
--    return elem.get_double();
  end;

  member function get_bool(pair_name varchar2) return boolean as
    elem pljson_value := get(pair_name);
  begin
    
    if elem is not null then
      return elem.get_bool();
    end if;
    return null;
    
--    return elem.get_bool();
  end;

  member function get_pljson(pair_name varchar2) return pljson as
    elem pljson_value := get(pair_name);
  begin
    
    if elem is not null then
      return treat(elem.object_or_array as pljson);
    end if;
    return null;
    
--    return treat(elem.object_or_array as pljson);
  end;

  member function get_pljson_list(pair_name varchar2) return pljson_list as
    elem pljson_value := get(pair_name);
  begin
    
    if elem is not null then
      return treat(elem.object_or_array as pljson_list);
    end if;
    return null;
    
--    return treat(elem.object_or_array as pljson_list);
  end;

  member function get(position pls_integer) return pljson_value as
  begin
    if (self.count >= position and position > 0) then
      return self.json_data(position);
    end if;
    return null; 
  end;

  member function index_of(pair_name varchar2) return number as
    indx pls_integer;
  begin
    indx := json_data.first;
    loop
      exit when indx is null;
      if (pair_name is null and json_data(indx).mapname is null) then return indx; end if;
      if (json_data(indx).mapname = pair_name) then return indx; end if;
      indx := json_data.next(indx);
    end loop;
    return -1;
  end;

  member function exist(pair_name varchar2) return boolean as
  begin
    return (get(pair_name) is not null);
  end;

  member function to_json_value return pljson_value as
  begin
    return pljson_value(self);
  end;

  member procedure check_duplicate(self in out nocopy pljson, v_set boolean) as
  begin
    if (v_set) then
      check_for_duplicate := 1;
    else
      check_for_duplicate := 0;
    end if;
  end;

  member procedure remove_duplicates(self in out nocopy pljson) as
  begin
    pljson_parser.remove_duplicates(self);
  end remove_duplicates;

--   output methods 
  member function to_char(spaces boolean default true, chars_per_line number default 0) return varchar2 as
  begin
    if(spaces is null) then
      return pljson_printer.pretty_print(self, line_length => chars_per_line);
    else
      return pljson_printer.pretty_print(self, spaces, line_length => chars_per_line);
    end if;
  end;

  member procedure to_clob(self in pljson, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true) as
  begin
    if(spaces is null) then
      pljson_printer.pretty_print(self, false, buf, line_length => chars_per_line, erase_clob => erase_clob);
    else
      pljson_printer.pretty_print(self, spaces, buf, line_length => chars_per_line, erase_clob => erase_clob);
    end if;
  end;

  member procedure print(self in pljson, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null) as --32512 is the real maximum in sqldeveloper
    my_clob clob;
  begin
    my_clob := empty_clob();
    dbms_lob.createtemporary(my_clob, true);
    pljson_printer.pretty_print(self, spaces, my_clob, case when (chars_per_line>32512) then 32512 else chars_per_line end);
    pljson_printer.dbms_output_clob(my_clob, pljson_printer.newline_char, jsonp);
    dbms_lob.freetemporary(my_clob);
  end;

  member procedure htp(self in pljson, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null) as
    my_clob clob;
  begin
    my_clob := empty_clob();
    dbms_lob.createtemporary(my_clob, true);
    pljson_printer.pretty_print(self, spaces, my_clob, chars_per_line);
    pljson_printer.htp_output_clob(my_clob, jsonp);
    dbms_lob.freetemporary(my_clob);
  end;

--   json path 
  member function path(json_path varchar2, base number default 1) return pljson_value as
  begin
    return pljson_ext.get_json_value(self, json_path, base);
  end path;

--   json path_put 
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem pljson_value, base number default 1) as
  begin
    pljson_ext.put(self, json_path, elem, base);
  end path_put;

  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem varchar2, base number default 1) as
  begin
    pljson_ext.put(self, json_path, elem, base);
  end path_put;

  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem clob, base number default 1) as
  begin
    pljson_ext.put(self, json_path, elem, base);
  end path_put;

  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem number, base number default 1) as
  begin
    if (elem is null) then
      pljson_ext.put(self, json_path, pljson_value(), base);
    else
      pljson_ext.put(self, json_path, elem, base);
    end if;
  end path_put;

--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem binary_double, base number default 1) as
  begin
    if (elem is null) then
      pljson_ext.put(self, json_path, pljson_value(), base);
    else
      pljson_ext.put(self, json_path, elem, base);
    end if;
  end path_put;

  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem boolean, base number default 1) as
  begin
    if (elem is null) then
      pljson_ext.put(self, json_path, pljson_value(), base);
    else
      pljson_ext.put(self, json_path, elem, base);
    end if;
  end path_put;

  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem pljson_list, base number default 1) as
  begin
    if (elem is null) then
      pljson_ext.put(self, json_path, pljson_value(), base);
    else
      pljson_ext.put(self, json_path, elem, base);
    end if;
  end path_put;

  member procedure path_put(self in out nocopy pljson, json_path varchar2, elem pljson, base number default 1) as
  begin
    if (elem is null) then
      pljson_ext.put(self, json_path, pljson_value(), base);
    else
      pljson_ext.put(self, json_path, elem, base);
    end if;
  end path_put;

  member procedure path_remove(self in out nocopy pljson, json_path varchar2, base number default 1) as
  begin
    pljson_ext.remove(self, json_path, base);
  end path_remove;

--   Thanks to Matt Nolan 
  member function get_keys return pljson_list as
    keys pljson_list;
    indx pls_integer;
  begin
    keys := pljson_list();
    indx := json_data.first;
    loop
      exit when indx is null;
      keys.append(json_data(indx).mapname);
      indx := json_data.next(indx);
    end loop;
    return keys;
  end;

  member function get_values return pljson_list as
    vals pljson_list := pljson_list();
  begin
    vals.list_data := self.json_data;
    return vals;
  end;
end;
/
create or replace package pljson_ac as
  --json type methods
  
  procedure object_remove(p_self in out nocopy pljson, pair_name varchar2);
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value pljson_value, position pls_integer default null);
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value varchar2, position pls_integer default null);
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value number, position pls_integer default null);
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value binary_double, position pls_integer default null);
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value boolean, position pls_integer default null);
  procedure object_check_duplicate(p_self in out nocopy pljson, v_set boolean);
  procedure object_remove_duplicates(p_self in out nocopy pljson);
  
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value pljson, position pls_integer default null);
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value pljson_list, position pls_integer default null);
  
  function object_count(p_self in pljson) return number;
  function object_get(p_self in pljson, pair_name varchar2) return pljson_value;
  function object_get(p_self in pljson, position pls_integer) return pljson_value;
  function object_index_of(p_self in pljson, pair_name varchar2) return number;
  function object_exist(p_self in pljson, pair_name varchar2) return boolean;
  
  function object_to_char(p_self in pljson, spaces boolean default true, chars_per_line number default 0) return varchar2;
  procedure object_to_clob(p_self in pljson, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true);
  procedure object_print(p_self in pljson, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null);
  procedure object_htp(p_self in pljson, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null);
  
  function object_to_json_value(p_self in pljson) return pljson_value;
  function object_path(p_self in pljson, json_path varchar2, base number default 1) return pljson_value;
  
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem pljson_value, base number default 1);
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem varchar2  , base number default 1);
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem number    , base number default 1);
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem binary_double, base number default 1);
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem boolean   , base number default 1);
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem pljson_list , base number default 1);
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem pljson      , base number default 1);
  
  procedure object_path_remove(p_self in out nocopy pljson, json_path varchar2, base number default 1);
  
  function object_get_values(p_self in pljson) return pljson_list;
  function object_get_keys(p_self in pljson) return pljson_list;
  
  --json_list type methods
  procedure array_append(p_self in out nocopy pljson_list, elem pljson_value, position pls_integer default null);
  procedure array_append(p_self in out nocopy pljson_list, elem varchar2, position pls_integer default null);
  procedure array_append(p_self in out nocopy pljson_list, elem number, position pls_integer default null);
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure array_append(p_self in out nocopy pljson_list, elem binary_double, position pls_integer default null);
  procedure array_append(p_self in out nocopy pljson_list, elem boolean, position pls_integer default null);
  procedure array_append(p_self in out nocopy pljson_list, elem pljson_list, position pls_integer default null);
  
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem pljson_value);
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem varchar2);
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem number);
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem binary_double);
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem boolean);
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem pljson_list);
  
  function array_count(p_self in pljson_list) return number;
  procedure array_remove(p_self in out nocopy pljson_list, position pls_integer);
  procedure array_remove_first(p_self in out nocopy pljson_list);
  procedure array_remove_last(p_self in out nocopy pljson_list);
  function array_get(p_self in pljson_list, position pls_integer) return pljson_value;
  function array_head(p_self in pljson_list) return pljson_value;
  function array_last(p_self in pljson_list) return pljson_value;
  function array_tail(p_self in pljson_list) return pljson_list;
  
  function array_to_char(p_self in pljson_list, spaces boolean default true, chars_per_line number default 0) return varchar2;
  procedure array_to_clob(p_self in pljson_list, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true);
  procedure array_print(p_self in pljson_list, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null);
  procedure array_htp(p_self in pljson_list, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null);
  
  function array_path(p_self in pljson_list, json_path varchar2, base number default 1) return pljson_value;
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem pljson_value, base number default 1);
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem varchar2  , base number default 1);
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem number    , base number default 1);
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem binary_double, base number default 1);
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem boolean   , base number default 1);
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem pljson_list , base number default 1);
  
  procedure array_path_remove(p_self in out nocopy pljson_list, json_path varchar2, base number default 1);
  
  function array_to_json_value(p_self in pljson_list) return pljson_value;
  
  --json_value
  
  
  function jv_get_type(p_self in pljson_value) return varchar2;
  function jv_get_string(p_self in pljson_value, max_byte_size number default null, max_char_size number default null) return varchar2;
  procedure jv_get_string(p_self in pljson_value, buf in out nocopy clob);
  function jv_get_number(p_self in pljson_value) return number;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function jv_get_double(p_self in pljson_value) return binary_double;
  function jv_get_bool(p_self in pljson_value) return boolean;
  function jv_get_null(p_self in pljson_value) return varchar2;
  
  function jv_is_object(p_self in pljson_value) return boolean;
  function jv_is_array(p_self in pljson_value) return boolean;
  function jv_is_string(p_self in pljson_value) return boolean;
  function jv_is_number(p_self in pljson_value) return boolean;
  function jv_is_bool(p_self in pljson_value) return boolean;
  function jv_is_null(p_self in pljson_value) return boolean;
  
  function jv_to_char(p_self in pljson_value, spaces boolean default true, chars_per_line number default 0) return varchar2;
  procedure jv_to_clob(p_self in pljson_value, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true);
  procedure jv_print(p_self in pljson_value, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null);
  procedure jv_htp(p_self in pljson_value, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null);
  
  function jv_value_of(p_self in pljson_value, max_byte_size number default null, max_char_size number default null) return varchar2;


end pljson_ac;
/

create or replace package body pljson_ac as
  procedure object_remove(p_self in out nocopy pljson, pair_name varchar2) as
  begin p_self.remove(pair_name); end;
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value pljson_value, position pls_integer default null) as
  begin p_self.put(pair_name, pair_value, position); end;
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value varchar2, position pls_integer default null) as
  begin p_self.put(pair_name, pair_value, position); end;
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value number, position pls_integer default null) as
  begin p_self.put(pair_name, pair_value, position); end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value binary_double, position pls_integer default null) as
  begin p_self.put(pair_name, pair_value, position); end;
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value boolean, position pls_integer default null) as
  begin p_self.put(pair_name, pair_value, position); end;
  procedure object_check_duplicate(p_self in out nocopy pljson, v_set boolean) as
  begin p_self.check_duplicate(v_set); end;
  procedure object_remove_duplicates(p_self in out nocopy pljson) as
  begin p_self.remove_duplicates; end;
  
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value pljson, position pls_integer default null) as
  begin p_self.put(pair_name, pair_value, position); end;
  procedure object_put(p_self in out nocopy pljson, pair_name varchar2, pair_value pljson_list, position pls_integer default null) as
  begin p_self.put(pair_name, pair_value, position); end;
  
  function object_count(p_self in pljson) return number as
  begin return p_self.count; end;
  function object_get(p_self in pljson, pair_name varchar2) return pljson_value as
  begin return p_self.get(pair_name); end;
  function object_get(p_self in pljson, position pls_integer) return pljson_value as
  begin return p_self.get(position); end;
  function object_index_of(p_self in pljson, pair_name varchar2) return number as
  begin return p_self.index_of(pair_name); end;
  function object_exist(p_self in pljson, pair_name varchar2) return boolean as
  begin return p_self.exist(pair_name); end;
  
  function object_to_char(p_self in pljson, spaces boolean default true, chars_per_line number default 0) return varchar2 as
  begin return p_self.to_char(spaces, chars_per_line); end;
  procedure object_to_clob(p_self in pljson, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true) as
  begin p_self.to_clob(buf, spaces, chars_per_line, erase_clob); end;
  procedure object_print(p_self in pljson, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null) as
  begin p_self.print(spaces, chars_per_line, jsonp); end;
  procedure object_htp(p_self in pljson, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null) as
  begin p_self.htp(spaces, chars_per_line, jsonp); end;
  
  function object_to_json_value(p_self in pljson) return pljson_value as
  begin return p_self.to_json_value; end;
  function object_path(p_self in pljson, json_path varchar2, base number default 1) return pljson_value as
  begin return p_self.path(json_path, base); end;
  
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem pljson_value, base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem varchar2  , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem number    , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem binary_double, base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem boolean   , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem pljson_list , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure object_path_put(p_self in out nocopy pljson, json_path varchar2, elem pljson      , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  
  procedure object_path_remove(p_self in out nocopy pljson, json_path varchar2, base number default 1) as
  begin p_self.path_remove(json_path, base); end;
  
  function object_get_values(p_self in pljson) return pljson_list as
  begin return p_self.get_values; end;
  function object_get_keys(p_self in pljson) return pljson_list as
  begin return p_self.get_keys; end;
  
  --json_list type
  procedure array_append(p_self in out nocopy pljson_list, elem pljson_value, position pls_integer default null) as
  begin p_self.append(elem, position); end;
  procedure array_append(p_self in out nocopy pljson_list, elem varchar2, position pls_integer default null) as
  begin p_self.append(elem, position); end;
  procedure array_append(p_self in out nocopy pljson_list, elem number, position pls_integer default null) as
  begin p_self.append(elem, position); end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure array_append(p_self in out nocopy pljson_list, elem binary_double, position pls_integer default null) as
  begin p_self.append(elem, position); end;
  procedure array_append(p_self in out nocopy pljson_list, elem boolean, position pls_integer default null) as
  begin p_self.append(elem, position); end;
  procedure array_append(p_self in out nocopy pljson_list, elem pljson_list, position pls_integer default null) as
  begin p_self.append(elem, position); end;
  
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem pljson_value) as
  begin p_self.replace(position, elem); end;
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem varchar2) as
  begin p_self.replace(position, elem); end;
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem number) as
  begin p_self.replace(position, elem); end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem binary_double) as
  begin p_self.replace(position, elem); end;
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem boolean) as
  begin p_self.replace(position, elem); end;
  procedure array_replace(p_self in out nocopy pljson_list, position pls_integer, elem pljson_list) as
  begin p_self.replace(position, elem); end;
  
  function array_count(p_self in pljson_list) return number as
  begin return p_self.count; end;
  procedure array_remove(p_self in out nocopy pljson_list, position pls_integer) as
  begin p_self.remove(position); end;
  procedure array_remove_first(p_self in out nocopy pljson_list) as
  begin p_self.remove_first; end;
  procedure array_remove_last(p_self in out nocopy pljson_list) as
  begin p_self.remove_last; end;
  function array_get(p_self in pljson_list, position pls_integer) return pljson_value as
  begin return p_self.get(position); end;
  function array_head(p_self in pljson_list) return pljson_value as
  begin return p_self.head; end;
  function array_last(p_self in pljson_list) return pljson_value as
  begin return p_self.last; end;
  function array_tail(p_self in pljson_list) return pljson_list as
  begin return p_self.tail; end;
  
  function array_to_char(p_self in pljson_list, spaces boolean default true, chars_per_line number default 0) return varchar2 as
  begin return p_self.to_char(spaces, chars_per_line); end;
  procedure array_to_clob(p_self in pljson_list, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true) as
  begin p_self.to_clob(buf, spaces, chars_per_line, erase_clob); end;
  procedure array_print(p_self in pljson_list, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null) as
  begin p_self.print(spaces, chars_per_line, jsonp); end;
  procedure array_htp(p_self in pljson_list, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null) as
  begin p_self.htp(spaces, chars_per_line, jsonp); end;
  
  function array_path(p_self in pljson_list, json_path varchar2, base number default 1) return pljson_value as
  begin return p_self.path(json_path, base); end;
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem pljson_value, base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem varchar2  , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem number    , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem binary_double, base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem boolean   , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  procedure array_path_put(p_self in out nocopy pljson_list, json_path varchar2, elem pljson_list , base number default 1) as
  begin p_self.path_put(json_path, elem, base); end;
  
  procedure array_path_remove(p_self in out nocopy pljson_list, json_path varchar2, base number default 1) as
  begin p_self.path_remove(json_path, base); end;
  
  function array_to_json_value(p_self in pljson_list) return pljson_value as
  begin return p_self.to_json_value; end;
  
  --json_value
  
  
  function jv_get_type(p_self in pljson_value) return varchar2 as
  begin return p_self.get_type; end;
  function jv_get_string(p_self in pljson_value, max_byte_size number default null, max_char_size number default null) return varchar2 as
  begin return p_self.get_string(max_byte_size, max_char_size); end;
  procedure jv_get_string(p_self in pljson_value, buf in out nocopy clob) as
  begin p_self.get_string(buf); end;
  function jv_get_number(p_self in pljson_value) return number as
  begin return p_self.get_number; end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function jv_get_double(p_self in pljson_value) return binary_double as
  begin return p_self.get_double; end;
  function jv_get_bool(p_self in pljson_value) return boolean as
  begin return p_self.get_bool; end;
  function jv_get_null(p_self in pljson_value) return varchar2 as
  begin return p_self.get_null; end;
  
  function jv_is_object(p_self in pljson_value) return boolean as
  begin return p_self.is_object; end;
  function jv_is_array(p_self in pljson_value) return boolean as
  begin return p_self.is_array; end;
  function jv_is_string(p_self in pljson_value) return boolean as
  begin return p_self.is_string; end;
  function jv_is_number(p_self in pljson_value) return boolean as
  begin return p_self.is_number; end;
  function jv_is_bool(p_self in pljson_value) return boolean as
  begin return p_self.is_bool; end;
  function jv_is_null(p_self in pljson_value) return boolean as
  begin return p_self.is_null; end;
  
  function jv_to_char(p_self in pljson_value, spaces boolean default true, chars_per_line number default 0) return varchar2 as
  begin return p_self.to_char(spaces, chars_per_line); end;
  procedure jv_to_clob(p_self in pljson_value, buf in out nocopy clob, spaces boolean default false, chars_per_line number default 0, erase_clob boolean default true) as
  begin p_self.to_clob(buf, spaces, chars_per_line, erase_clob); end;
  procedure jv_print(p_self in pljson_value, spaces boolean default true, chars_per_line number default 8192, jsonp varchar2 default null) as
  begin p_self.print(spaces, chars_per_line, jsonp); end;
  procedure jv_htp(p_self in pljson_value, spaces boolean default false, chars_per_line number default 0, jsonp varchar2 default null) as
  begin p_self.htp(spaces, chars_per_line, jsonp); end;
  
  function jv_value_of(p_self in pljson_value, max_byte_size number default null, max_char_size number default null) return varchar2 as
  begin return p_self.value_of(max_byte_size, max_char_size); end;

end pljson_ac;
/
create or replace package pljson_dyn authid current_user as
-- 
--  Copyright (c) 2010 Jonas Krogsboell
--
--  Permission is hereby granted, free of charge, to any person obtaining a copy
--  of this software and associated documentation files (the "Software"), to deal
--  in the Software without restriction, including without limitation the rights
--  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
--  copies of the Software, and to permit persons to whom the Software is
--  furnished to do so, subject to the following conditions:
--
--  The above copyright notice and this permission notice shall be included in
--  all copies or substantial portions of the Software.
--
--  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
--  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
--  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
--  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
--  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
--  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
--  THE SOFTWARE.
--  

  null_as_empty_string   boolean not null := true;  --varchar2
  include_dates          boolean not null := true;  --date
  include_clobs          boolean not null := true;
  include_blobs          boolean not null := false;
  include_arrays         boolean not null := true;  -- pljson_varray or pljson_narray

--   list with objects 
  function executeList(stmt varchar2, bindvar pljson default null, cur_num number default null, bindvardateformats pljson default null) return pljson_list;

--   object with lists 
  function executeObject(stmt varchar2, bindvar pljson default null, cur_num number default null) return pljson;


--   usage example:
--   * declare
--   *   res json_list;
--   * begin
--   *   res := json_dyn.executeList(
--   *            'select :bindme as one, :lala as two from dual where dummy in :arraybind',
--   *            json('{bindme:"4", lala:123, arraybind:[1, 2, 3, "X"]}')
--   *          );
--   *   res.print;
--   * end;
--   

-- --11g functions
--  function executeList(stmt in out sys_refcursor) return json_list;
--  function executeObject(stmt in out sys_refcursor) return json;
--
end pljson_dyn;
/

create or replace package body pljson_dyn as
--
--  -- 11gR2
--  function executeList(stmt in out sys_refcursor) return json_list as
--    l_cur number;
--  begin
--    l_cur := dbms_sql.to_cursor_number(stmt);
--    return json_dyn.executeList(null, null, l_cur);
--  end;
--
--  -- 11gR2
--  function executeObject(stmt in out sys_refcursor) return json as
--    l_cur number;
--  begin
--    l_cur := dbms_sql.to_cursor_number(stmt);
--    return json_dyn.executeObject(null, null, l_cur);
--  end;
--

  procedure bind_json(l_cur number, bindvar pljson, bindvardateformats pljson default null) as
    keylist pljson_list := bindvar.get_keys();
  begin
    for i in 1 .. keylist.count loop
      if (bindvar.get(i).is_number()) then
        dbms_sql.bind_variable(l_cur, ':'||keylist.get(i).get_string(), bindvar.get(i).get_number());
      elsif (bindvar.get(i).is_array()) then
        declare
          v_bind dbms_sql.varchar2_table;
          v_arr  pljson_list := pljson_list(bindvar.get(i));
        begin
          for j in 1 .. v_arr.count loop
            v_bind(j) := v_arr.get(j).value_of();
          end loop;
          dbms_sql.bind_array(l_cur, ':'||keylist.get(i).get_string(), v_bind);
        end;
      else
        if bindvardateformats is not null then
            if bindvardateformats.exist(keylist.get(i).get_string()) then
                dbms_sql.bind_variable(l_cur, ':'||keylist.get(i).get_string(), to_date(bindvar.get(i).value_of(), bindvardateformats.get(keylist.get(i).get_string()).get_string() ));
            else
                dbms_sql.bind_variable(l_cur, ':'||keylist.get(i).get_string(), bindvar.get(i).value_of());
            end if;
        else
            dbms_sql.bind_variable(l_cur, ':'||keylist.get(i).get_string(), bindvar.get(i).value_of());
        end if;
      end if;
    end loop;
  end bind_json;

--   list with objects 
  function executeList(stmt varchar2, bindvar pljson, cur_num number, bindvardateformats pljson default null) return pljson_list as
    l_cur number;
    --lshuo注释
    --由于10g不支持DESC_TAB3 ，且返回数组对象，pljson_narray，pljson_varray
    l_dtbl 
    $if dbms_DB_version.version>10 $then
	dbms_sql.desc_tab3;
    $else
	dbms_sql.desc_tab2;
    $end
    l_cnt number;
    l_status number;
    l_val varchar2(4000);
    outer_list pljson_list := pljson_list();
    inner_obj pljson;
    conv number;
    read_date date;
    read_clob clob;
    read_blob blob;
    col_type number;
    read_varray pljson_varray;
    read_narray pljson_narray;
  begin
    if (cur_num is not null) then
      l_cur := cur_num;
    else
      l_cur := dbms_sql.open_cursor;
      dbms_sql.parse(l_cur, stmt, dbms_sql.native);
      if (bindvar is not null) then bind_json(l_cur, bindvar, bindvardateformats); end if;
    end if;
--     E.I.Sarmas (github.com/dsnz)   2018-05-01   handling of varray, narray in select 
    $if dbms_DB_version.version>10 $then
	dbms_sql.describe_columns3(l_cur, l_cnt, l_dtbl);
    $else
        dbms_sql.describe_columns2(l_cur, l_cnt, l_dtbl);
    $end
    for i in 1..l_cnt loop
      col_type := l_dtbl(i).col_type;
      --dbms_output.put_line(col_type);
      if (col_type = 12) then
        dbms_sql.define_column(l_cur, i, read_date);
      elsif (col_type = 112) then
        dbms_sql.define_column(l_cur, i, read_clob);
      elsif (col_type = 113) then
        dbms_sql.define_column(l_cur, i, read_blob);
      elsif (col_type in (1, 2, 96)) then
        dbms_sql.define_column(l_cur, i, l_val, 4000);
      --lshuo注释
      --由于10g不支持DESC_TAB3 ，且无法返回数组对象，pljson_narray，pljson_varray
      $if dbms_DB_version.version>10 $then
	-- E.I.Sarmas (github.com/dsnz)   2018-05-01   handling of pljson_varray in select 
	elsif (col_type = 109 and l_dtbl(i).col_type_name = 'PLJSON_VARRAY') then
	  dbms_sql.define_column(l_cur, i, read_varray);
	-- E.I.Sarmas (github.com/dsnz)   2018-05-01   handling of pljson_narray in select 
	elsif (col_type = 109 and l_dtbl(i).col_type_name = 'PLJSON_NARRAY') then
	  dbms_sql.define_column(l_cur, i, read_narray);
      $end
--       E.I.Sarmas (github.com/dsnz)   2018-05-01   record unhandled col_type 
      else
        dbms_output.put_line('unhandled col_type =' || col_type);
      end if;
    end loop;

    if (cur_num is null) then l_status := dbms_sql.execute(l_cur); end if;

    --loop through rows
    while ( dbms_sql.fetch_rows(l_cur) > 0 ) loop
      inner_obj := pljson(); --init for each row
      inner_obj.check_for_duplicate := 0;
      --loop through columns
      for i in 1..l_cnt loop
        case true
        --handling string types
        when l_dtbl(i).col_type in (1, 96) then -- varchar2
          dbms_sql.column_value(l_cur, i, l_val);
          if (l_val is null) then
            if (null_as_empty_string) then
              inner_obj.put(l_dtbl(i).col_name, ''); --treat as emptystring 
            else
              inner_obj.put(l_dtbl(i).col_name, pljson_value()); --null
            end if;
          else
            inner_obj.put(l_dtbl(i).col_name, pljson_value(l_val)); --null
          end if;
          --dbms_output.put_line(l_dtbl(i).col_name||' --> '||l_val||'varchar2' ||l_dtbl(i).col_type);
        --handling number types
        when l_dtbl(i).col_type = 2 then -- number
          dbms_sql.column_value(l_cur, i, l_val);
          conv := l_val;
          inner_obj.put(l_dtbl(i).col_name, conv);
          -- dbms_output.put_line(l_dtbl(i).col_name||' --> '||l_val||'number ' ||l_dtbl(i).col_type);
        when l_dtbl(i).col_type = 12 then -- date
          if (include_dates) then
            dbms_sql.column_value(l_cur, i, read_date);
            inner_obj.put(l_dtbl(i).col_name, pljson_ext.to_json_value(read_date));
          end if;
          --dbms_output.put_line(l_dtbl(i).col_name||' --> '||l_val||'date ' ||l_dtbl(i).col_type);
        when l_dtbl(i).col_type = 112 then --clob
          if (include_clobs) then
            dbms_sql.column_value(l_cur, i, read_clob);
            inner_obj.put(l_dtbl(i).col_name, pljson_value(read_clob));
          end if;
        when l_dtbl(i).col_type = 113 then --blob
          if (include_blobs) then
            dbms_sql.column_value(l_cur, i, read_blob);
            if (dbms_lob.getlength(read_blob) > 0) then
              inner_obj.put(l_dtbl(i).col_name, pljson_ext.encode(read_blob));
            else
              inner_obj.put(l_dtbl(i).col_name, pljson_value());
            end if;
          end if;
        --lshuo注释
        --由于10g不支持DESC_TAB3 ，且无法返回数组对象，pljson_narray，pljson_varray
        $if dbms_DB_version.version>10 $then
		when l_dtbl(i).col_type = 109 and l_dtbl(i).col_type_name = 'PLJSON_VARRAY' then
		  if (include_arrays) then
		    dbms_sql.column_value(l_cur, i, read_varray);
		    inner_obj.put(l_dtbl(i).col_name, pljson_list(read_varray));
		  end if;
		-- E.I.Sarmas (github.com/dsnz)   2018-05-01   handling of pljson_narray in select 
		when l_dtbl(i).col_type = 109 and l_dtbl(i).col_type_name = 'PLJSON_NARRAY' then
		  if (include_arrays) then
		    dbms_sql.column_value(l_cur, i, read_narray);
		    inner_obj.put(l_dtbl(i).col_name, pljson_list(read_narray));
		  end if;
        $end
        else null; --discard other types
        end case;
      end loop;
      inner_obj.check_for_duplicate := 1;
      outer_list.append(inner_obj.to_json_value);
    end loop;
    dbms_sql.close_cursor(l_cur);
    return outer_list;
  end executeList;

--   object with lists 
  function executeObject(stmt varchar2, bindvar pljson, cur_num number) return pljson as
    l_cur number;
    l_dtbl dbms_sql.desc_tab;
    l_cnt number;
    l_status number;
    l_val varchar2(4000);
    inner_list_names pljson_list := pljson_list();
    inner_list_data pljson_list := pljson_list();
    data_list pljson_list;
    outer_obj pljson := pljson();
    conv number;
    read_date date;
    read_clob clob;
    read_blob blob;
    col_type number;
  begin
    if (cur_num is not null) then
      l_cur := cur_num;
    else
      l_cur := dbms_sql.open_cursor;
      dbms_sql.parse(l_cur, stmt, dbms_sql.native);
      if (bindvar is not null) then bind_json(l_cur, bindvar); end if;
    end if;
    dbms_sql.describe_columns(l_cur, l_cnt, l_dtbl);
    for i in 1..l_cnt loop
      col_type := l_dtbl(i).col_type;
      if (col_type = 12) then
        dbms_sql.define_column(l_cur, i, read_date);
      elsif (col_type = 112) then
        dbms_sql.define_column(l_cur, i, read_clob);
      elsif (col_type = 113) then
        dbms_sql.define_column(l_cur, i, read_blob);
      elsif (col_type in (1, 2, 96)) then
        dbms_sql.define_column(l_cur, i, l_val, 4000);
      end if;
    end loop;
    if (cur_num is null) then l_status := dbms_sql.execute(l_cur); end if;

    --build up name_list
    for i in 1..l_cnt loop
      case l_dtbl(i).col_type
        when 1 then inner_list_names.append(l_dtbl(i).col_name);
        when 96 then inner_list_names.append(l_dtbl(i).col_name);
        when 2 then inner_list_names.append(l_dtbl(i).col_name);
        when 12 then if (include_dates) then inner_list_names.append(l_dtbl(i).col_name); end if;
        when 112 then if (include_clobs) then inner_list_names.append(l_dtbl(i).col_name); end if;
        when 113 then if (include_blobs) then inner_list_names.append(l_dtbl(i).col_name); end if;
        else null;
      end case;
    end loop;

    --loop through rows
    while ( dbms_sql.fetch_rows(l_cur) > 0 ) loop
      data_list := pljson_list();
      --loop through columns
      for i in 1..l_cnt loop
        case true
        --handling string types
        when l_dtbl(i).col_type in (1, 96) then -- varchar2
          dbms_sql.column_value(l_cur, i, l_val);
          if (l_val is null) then
            if (null_as_empty_string) then
              data_list.append(''); --treat as emptystring 
            else
              data_list.append(pljson_value()); --null
            end if;
          else
            data_list.append(pljson_value(l_val)); --null
          end if;
          --dbms_output.put_line(l_dtbl(i).col_name||' --> '||l_val||'varchar2' ||l_dtbl(i).col_type);
        --handling number types
        when l_dtbl(i).col_type = 2 then -- number
          dbms_sql.column_value(l_cur, i, l_val);
          conv := l_val;
          data_list.append(conv);
          -- dbms_output.put_line(l_dtbl(i).col_name||' --> '||l_val||'number ' ||l_dtbl(i).col_type);
        when l_dtbl(i).col_type = 12 then -- date
          if (include_dates) then
            dbms_sql.column_value(l_cur, i, read_date);
            data_list.append(pljson_ext.to_json_value(read_date));
          end if;
          --dbms_output.put_line(l_dtbl(i).col_name||' --> '||l_val||'date ' ||l_dtbl(i).col_type);
        when l_dtbl(i).col_type = 112 then --clob
          if (include_clobs) then
            dbms_sql.column_value(l_cur, i, read_clob);
            data_list.append(pljson_value(read_clob));
          end if;
        when l_dtbl(i).col_type = 113 then --blob
          if (include_blobs) then
            dbms_sql.column_value(l_cur, i, read_blob);
            if (dbms_lob.getlength(read_blob) > 0) then
              data_list.append(pljson_ext.encode(read_blob));
            else
              data_list.append(pljson_value());
            end if;
          end if;
        else null; --discard other types
        end case;
      end loop;
      inner_list_data.append(data_list);
    end loop;

    outer_obj.put('names', inner_list_names.to_json_value);
    outer_obj.put('data', inner_list_data.to_json_value);
    dbms_sql.close_cursor(l_cur);
    return outer_obj;
  end executeObject;

end pljson_dyn;
/

create or replace package pljson_ml as
--  
--  Copyright (c) 2010 Jonas Krogsboell
--
--  Permission is hereby granted, free of charge, to any person obtaining a copy
--  of this software and associated documentation files (the "Software"), to deal
--  in the Software without restriction, including without limitation the rights
--  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
--  copies of the Software, and to permit persons to whom the Software is
--  furnished to do so, subject to the following conditions:
--
--  The above copyright notice and this permission notice shall be included in
--  all copies or substantial portions of the Software.
--
--  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
--  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
--  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
--  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
--  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
--  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
--  THE SOFTWARE.
--  
  
--   This package contains extra methods to lookup types and
--     an easy way of adding date values in json - without changing the structure 

  jsonml_stylesheet xmltype := null;

  function xml2json(xml in xmltype) return pljson_list;
  function xmlstr2json(xmlstr in varchar2) return pljson_list;

end pljson_ml;
/

create or replace package body pljson_ml as
  function get_jsonml_stylesheet return xmltype;

  function xml2json(xml in xmltype) return pljson_list as
    l_json        xmltype;
    l_returnvalue clob;
  begin
    l_json := xml.transform (get_jsonml_stylesheet);
    l_returnvalue := l_json.getclobval();
    l_returnvalue := dbms_xmlgen.convert (l_returnvalue, dbms_xmlgen.entity_decode);
    --dbms_output.put_line(l_returnvalue);
    return pljson_list(l_returnvalue);
  end xml2json;

  function xmlstr2json(xmlstr in varchar2) return pljson_list as
  begin
    return xml2json(xmltype(xmlstr));
  end xmlstr2json;

  function get_jsonml_stylesheet return xmltype as
  begin
    if (jsonml_stylesheet is null) then
    jsonml_stylesheet := xmltype('<?xml version="1.0" encoding="UTF-8"?>
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
    end if;
    return jsonml_stylesheet;
  end get_jsonml_stylesheet;

end pljson_ml;
/
create or replace package pljson_xml as
--  
--  Copyright (c) 2010 Jonas Krogsboell
--
--  Permission is hereby granted, free of charge, to any person obtaining a copy
--  of this software and associated documentation files (the "Software"), to deal
--  in the Software without restriction, including without limitation the rights
--  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
--  copies of the Software, and to permit persons to whom the Software is
--  furnished to do so, subject to the following conditions:
--
--  The above copyright notice and this permission notice shall be included in
--  all copies or substantial portions of the Software.
--
--  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
--  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
--  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
--  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
--  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
--  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
--  THE SOFTWARE.
--  

--  
--  declare
--    obj json := json('{a:1, b:[2, 3, 4], c:true}');
--    x xmltype;
--  begin
--    obj.print;
--    x := json_xml.json_to_xml(obj);
--    dbms_output.put_line(x.getclobval());
--  end;
--  

  function json_to_xml(obj pljson, tagname varchar2 default 'root') return xmltype;

end pljson_xml;
/

create or replace package body pljson_xml as

  function escapeStr(str varchar2) return varchar2 as
    buf varchar2(32767) := '';
    ch varchar2(4);
  begin
    for i in 1 .. length(str) loop
      ch := substr(str, i, 1);
      case ch
      when chr(38)  then buf := buf ||chr(38) ||'amp;';
      when '<' then buf := buf ||chr(38) ||'lt;';
      when '>' then buf := buf ||chr(38) ||'gt;';
      when '"' then buf := buf ||chr(38) ||'quot;';
      else buf := buf || ch;
      end case;
    end loop;
    return buf;
  end escapeStr;

-- Clob methods from printer 
  procedure add_to_clob(buf_lob in out nocopy clob, buf_str in out nocopy varchar2, str varchar2) as
  begin
    if (length(str) > 32767 - length(buf_str)) then
      dbms_lob.append(buf_lob, buf_str);
      buf_str := str;
    else
      buf_str := buf_str || str;
    end if;
  end add_to_clob;
  
  procedure flush_clob(buf_lob in out nocopy clob, buf_str in out nocopy varchar2) as
  begin
    dbms_lob.append(buf_lob, buf_str);
  end flush_clob;
  
  procedure toString(obj pljson_value, tagname in varchar2, xmlstr in out nocopy clob, xmlbuf in out nocopy varchar2) as
    v_obj pljson;
    v_list pljson_list;
    
    v_keys pljson_list;
    v_value pljson_value;
    key_str varchar2(4000);
  begin
    if (obj.is_object()) then
      add_to_clob(xmlstr, xmlbuf, '<' || tagname || '>');
      v_obj := pljson(obj);
      
      v_keys := v_obj.get_keys();
      for i in 1 .. v_keys.count loop
        v_value := v_obj.get(i);
        key_str := v_keys.get(i).get_string();
        
        if (key_str = 'content') then
          if (v_value.is_array()) then
            declare
              v_l pljson_list := pljson_list(v_value);
            begin
              for j in 1 .. v_l.count loop
                if (j > 1) then add_to_clob(xmlstr, xmlbuf, chr(13)||chr(10)); end if;
                add_to_clob(xmlstr, xmlbuf, escapeStr(v_l.get(j).to_char()));
              end loop;
            end;
          else
            add_to_clob(xmlstr, xmlbuf, escapeStr(v_value.to_char()));
          end if;
        elsif (v_value.is_array()) then
          declare
            v_l pljson_list := pljson_list(v_value);
          begin
            for j in 1 .. v_l.count loop
              v_value := v_l.get(j);
              if (v_value.is_array()) then
                add_to_clob(xmlstr, xmlbuf, '<' || key_str || '>');
                add_to_clob(xmlstr, xmlbuf, escapeStr(v_value.to_char()));
                add_to_clob(xmlstr, xmlbuf, '</' || key_str || '>');
              else
                toString(v_value, key_str, xmlstr, xmlbuf);
              end if;
            end loop;
          end;
        elsif (v_value.is_null() or (v_value.is_string() and v_value.get_string() = '')) then
          add_to_clob(xmlstr, xmlbuf, '<' || key_str || '/>');
        else
          toString(v_value, key_str, xmlstr, xmlbuf);
        end if;
      end loop;
      
      add_to_clob(xmlstr, xmlbuf, '</' || tagname || '>');
    elsif (obj.is_array()) then
      v_list := pljson_list(obj);
      for i in 1 .. v_list.count loop
        v_value := v_list.get(i);
        toString(v_value, nvl(tagname, 'array'), xmlstr, xmlbuf);
      end loop;
    else
      add_to_clob(xmlstr, xmlbuf, '<' || tagname || '>'||case when obj.value_of() is  not null then escapeStr(obj.value_of()) end ||'</' || tagname || '>');
    end if;
  end toString;
  
  function json_to_xml(obj pljson, tagname varchar2 default 'root') return xmltype as
    xmlstr clob := empty_clob();
    xmlbuf varchar2(32767) := '';
    returnValue xmltype;
  begin
    dbms_lob.createtemporary(xmlstr, true);
    
    toString(obj.to_json_value, tagname, xmlstr, xmlbuf);
    
    flush_clob(xmlstr, xmlbuf);
    returnValue := xmltype('<?xml version="1.0"?>'||xmlstr);
    dbms_lob.freetemporary(xmlstr);
    return returnValue;
  end;

end pljson_xml;
/

create or replace package pljson_util_pkg authid current_user as

--  
--
--  Purpose:    JSON utilities for PL/SQL
--  see http://ora-00001.blogspot.com/
--  
--  Remarks:
--  
--  Who     Date        Description
--  ------  ----------  -------------------------------------
--  MBR     30.01.2010  Created
--  JKR     01.05.2010  Edited to fit in PL/JSON
--  JKR     19.01.2011  Newest stylesheet + bugfix handling
--  
--  

  -- generate JSON from REF Cursor
  function ref_cursor_to_json (p_ref_cursor in sys_refcursor,
                               p_max_rows in number := null,
                               p_skip_rows in number := null) return pljson_list;

  -- generate JSON from SQL statement
  function sql_to_json (p_sql in varchar2,
                        p_max_rows in number := null,
                        p_skip_rows in number := null) return pljson_list;


end pljson_util_pkg;
/

create or replace package body pljson_util_pkg as
  scanner_exception exception;
  pragma exception_init(scanner_exception, -20100);
  parser_exception exception;
  pragma exception_init(parser_exception, -20101);

--  
--
--  Purpose:    JSON utilities for PL/SQL
--
--  Remarks:
--
--  Who     Date        Description
--  ------  ----------  -------------------------------------
--  MBR     30.01.2010  Created
--
--  


  g_json_null_object             constant varchar2(20) := '{ }';


  Function Get_Xml_To_Json_Stylesheet Return Xmltype As
  Begin
  
--    
--    
--    Purpose:    return XSLT stylesheet for XML to JSON transformation
--    
--    Remarks:    see http://code.google.com/p/xml2json-xslt/
--    
--    Who     Date        Description
--    ------  ----------  -------------------------------------
--    MBR     30.01.2010  Created
--    MBR     30.01.2010  Added fix for nulls
--    
--    
  --lshuo注释
  --去掉去q转义。去掉开头^与结尾^,单引号调整为双个单引号，ANd符号替换为Chr(38)
  --通过以上修改，保证可以PLSQL直接执行
    Return Xmltype('<?xml version="1.0" encoding="UTF-8"?>
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
      <xsl:when test="contains($s,''\'')">
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="concat(substring-before($s,''\''),''\\'')"/>
        </xsl:call-template>
        <xsl:call-template name="escape-bs-string">
          <xsl:with-param name="s" select="substring-after($s,''\'')"/>
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
      <xsl:when test="contains($s,''' || Chr(38) ||
                   'quot;'')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,''' || Chr(38) ||
                   'quot;''),''\' || Chr(38) || 'quot;'')"/>
        </xsl:call-template>
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="substring-after($s,''' || Chr(38) ||
                   'quot;'')"/>
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
       or double quote here, because they don''t replace characters (' || Chr(38) ||
                   '#x0; becomes \t), but they prefix
       characters (\ becomes \\). Besides, backslash should be seperate anyway, because it should be
       processed first. This function can''t do that. -->
  <xsl:template name="encode-string">
    <xsl:param name="s"/>
    <xsl:choose>
      <!-- tab -->
      <xsl:when test="contains($s,''' || Chr(38) ||
                   '#x9;'')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,''' || Chr(38) ||
                   '#x9;''),''\t'',substring-after($s,''' || Chr(38) || '#x9;''))"/>
        </xsl:call-template>
      </xsl:when>
      <!-- line feed -->
      <xsl:when test="contains($s,''' || Chr(38) ||
                   '#xA;'')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,''' || Chr(38) ||
                   '#xA;''),''\n'',substring-after($s,''' || Chr(38) || '#xA;''))"/>
        </xsl:call-template>
      </xsl:when>
      <!-- carriage return -->
      <xsl:when test="contains($s,''' || Chr(38) ||
                   '#xD;'')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,''' || Chr(38) ||
                   '#xD;''),''\r'',substring-after($s,''' || Chr(38) ||
                   '#xD;''))"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise><xsl:value-of select="$s"/></xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- number (no support for javascript mantissa) -->
  <xsl:template match="text()[not(string(number())=''NaN'' or
                      (starts-with(.,''0'' ) and . != ''0'' and
not(starts-with(.,''0.'' ))) or
                      (starts-with(.,''-0'' ) and . != ''-0'' and
not(starts-with(.,''-0.'' )))
                      )]">
    <xsl:value-of select="."/>
  </xsl:template>

  <!-- boolean, case-insensitive -->
  <xsl:template match="text()[translate(.,''TRUE'',''true'')=''true'']">true</xsl:template>
  <xsl:template match="text()[translate(.,''FALSE'',''false'')=''false'']">false</xsl:template>

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
  <xsl:template match="*[count(../*[name(../*)=name(.)])=count(../*) and count(../*)' || Chr(38) ||
                   'gt;1]">
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

</xsl:stylesheet>');
  
  End Get_Xml_To_Json_Stylesheet;


function ref_cursor_to_json (p_ref_cursor in sys_refcursor,
                             p_max_rows in number := null,
                             p_skip_rows in number := null) return pljson_list
as
  l_ctx         dbms_xmlgen.ctxhandle;
  l_num_rows    pls_integer;
  l_xml         xmltype;
  l_json        xmltype;
  l_returnvalue clob;
begin

--  
--
--  Purpose:    generate JSON from REF Cursor
--
--  Remarks:
--
--  Who     Date        Description
--  ------  ----------  -------------------------------------
--  MBR     30.01.2010  Created
--  JKR     01.05.2010  Edited to fit in PL/JSON
--
--  

  l_ctx := dbms_xmlgen.newcontext (p_ref_cursor);
  
  dbms_xmlgen.setnullhandling (l_ctx, dbms_xmlgen.empty_tag);
  
  -- for pagination
  
  if p_max_rows is not null then
    dbms_xmlgen.setmaxrows (l_ctx, p_max_rows);
  end if;
  
  if p_skip_rows is not null then
    dbms_xmlgen.setskiprows (l_ctx, p_skip_rows);
  end if;
  
  -- get the XML content
  l_xml := dbms_xmlgen.getxmltype (l_ctx, dbms_xmlgen.none);
  
  l_num_rows := dbms_xmlgen.getnumrowsprocessed (l_ctx);
  
  dbms_xmlgen.closecontext (l_ctx);
  
  close p_ref_cursor;
  
  if l_num_rows > 0 then
    -- perform the XSL transformation
    --lshuo注释
    --不用XMLTypr get_xml_to_json_stylesheet直接返回XMTYPE
    l_json := l_xml.transform (get_xml_to_json_stylesheet);
    l_returnvalue := l_json.getclobval();
  else
    l_returnvalue := g_json_null_object;
  end if;
  
  l_returnvalue := dbms_xmlgen.convert (l_returnvalue, dbms_xmlgen.entity_decode);
  
  if(l_num_rows = 0) then
    return pljson_list();
  else
    if(l_num_rows = 1) then
      declare ret pljson_list := pljson_list();
      begin
        ret.append(
          pljson(
            pljson(l_returnvalue).get('ROWSET')
          ).get('ROW')
        );
        return ret;
      end;
    else
      return pljson_list(pljson(l_returnvalue).get('ROWSET'));
    end if;
  end if;

exception
  when scanner_exception then
    dbms_output.put('Scanner problem with the following input: ');
    dbms_output.put_line(l_returnvalue);
    raise;
  when parser_exception then
    dbms_output.put('Parser problem with the following input: ');
    dbms_output.put_line(l_returnvalue);
    raise;
  when others then raise;
end ref_cursor_to_json;

function sql_to_json (p_sql in varchar2,
                      p_max_rows in number := null,
                      p_skip_rows in number := null) return pljson_list
as
  v_cur sys_refcursor;
begin
  open v_cur for p_sql;
  return ref_cursor_to_json(v_cur, p_max_rows, p_skip_rows);

end sql_to_json;


end pljson_util_pkg;
/
create or replace package pljson_helper as
 
  -- Recursive merge
  -- Courtesy of Matt Nolan - edited by Jonas Krogsboell
  function merge(p_a_json pljson, p_b_json pljson) return pljson;
  
  -- Join two lists
 
  function join(p_a_list pljson_list, p_b_list pljson_list) return pljson_list;
  
  -- keep only specific keys in json object
 
  function keep(p_json pljson, p_keys pljson_list) return pljson;
  
  -- remove specific keys in json object
 
  function remove(p_json pljson, p_keys pljson_list) return pljson;
  
  --equals
  function equals(p_v1 pljson_value, p_v2 pljson_value, exact boolean default true) return boolean;
  function equals(p_v1 pljson_value, p_v2 pljson, exact boolean default true) return boolean;
  function equals(p_v1 pljson_value, p_v2 pljson_list, exact boolean default true) return boolean;
  function equals(p_v1 pljson_value, p_v2 number) return boolean;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function equals(p_v1 pljson_value, p_v2 binary_double) return boolean;  
  function equals(p_v1 pljson_value, p_v2 varchar2) return boolean;
  function equals(p_v1 pljson_value, p_v2 boolean) return boolean;
  function equals(p_v1 pljson_value, p_v2 clob) return boolean;
  function equals(p_v1 pljson, p_v2 pljson, exact boolean default true) return boolean;
  function equals(p_v1 pljson_list, p_v2 pljson_list, exact boolean default true) return boolean;
  
  --contains json, json_value
  --contains json_list, json_value
  function contains(p_v1 pljson, p_v2 pljson_value, exact boolean default false) return boolean;
  function contains(p_v1 pljson, p_v2 pljson, exact boolean default false) return boolean;
  function contains(p_v1 pljson, p_v2 pljson_list, exact boolean default false) return boolean;
  function contains(p_v1 pljson, p_v2 number, exact boolean default false) return boolean;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function contains(p_v1 pljson, p_v2 binary_double, exact boolean default false) return boolean;
  function contains(p_v1 pljson, p_v2 varchar2, exact boolean default false) return boolean;
  function contains(p_v1 pljson, p_v2 boolean, exact boolean default false) return boolean;
  function contains(p_v1 pljson, p_v2 clob, exact boolean default false) return boolean;
  
  function contains(p_v1 pljson_list, p_v2 pljson_value, exact boolean default false) return boolean;
  function contains(p_v1 pljson_list, p_v2 pljson, exact boolean default false) return boolean;
  function contains(p_v1 pljson_list, p_v2 pljson_list, exact boolean default false) return boolean;
  function contains(p_v1 pljson_list, p_v2 number, exact boolean default false) return boolean;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function contains(p_v1 pljson_list, p_v2 binary_double, exact boolean default false) return boolean;
  function contains(p_v1 pljson_list, p_v2 varchar2, exact boolean default false) return boolean;
  function contains(p_v1 pljson_list, p_v2 boolean, exact boolean default false) return boolean;
  function contains(p_v1 pljson_list, p_v2 clob, exact boolean default false) return boolean;

end pljson_helper;
/

create or replace package body pljson_helper as
  
  --recursive merge
  function merge(p_a_json pljson, p_b_json pljson) return pljson as
    l_json    pljson;
    l_jv      pljson_value;
    l_indx    number;
    l_recursive pljson_value;
  begin
    --
    -- Initialize our return object
    --
    l_json := p_a_json;
    
    -- loop through p_b_json
    l_indx := p_b_json.json_data.first;
    loop
      exit when l_indx is null;
      l_jv   := p_b_json.json_data(l_indx);
      if (l_jv.is_object) then
        --recursive
        l_recursive := l_json.get(l_jv.mapname);
        if (l_recursive is not null and l_recursive.is_object) then
          l_json.put(l_jv.mapname, merge(pljson(l_recursive), pljson(l_jv)));
        else
          l_json.put(l_jv.mapname, l_jv);
        end if;
      else
        l_json.put(l_jv.mapname, l_jv);
      end if;
      
      --increment
      l_indx := p_b_json.json_data.next(l_indx);
    end loop;
    
    return l_json;
  
  end merge;
  
  -- join two lists
  function join(p_a_list pljson_list, p_b_list pljson_list) return pljson_list as
    l_json_list pljson_list := p_a_list;
  begin
    for indx in 1 .. p_b_list.count loop
      l_json_list.append(p_b_list.get(indx));
    end loop;
    
    return l_json_list;
  
  end join;
  
  -- keep keys.
  function keep(p_json pljson, p_keys pljson_list) return pljson as
    l_json pljson := pljson();
    mapname varchar2(4000);
  begin
    for i in 1 .. p_keys.count loop
      mapname := p_keys.get(i).get_string();
      if (p_json.exist(mapname)) then
        l_json.put(mapname, p_json.get(mapname));
      end if;
    end loop;
    
    return l_json;
  end keep;
  
  -- drop keys.
  function remove(p_json pljson, p_keys pljson_list) return pljson as
    l_json pljson := p_json;
  begin
    for i in 1 .. p_keys.count loop
      l_json.remove(p_keys.get(i).get_string());
    end loop;
    
    return l_json;
  end remove;
  
  --equals functions
  
  function equals(p_v1 pljson_value, p_v2 number) return boolean as
  begin
    if (p_v2 is null) then
      return p_v1.is_null;
    end if;
    
    if (not p_v1.is_number) then
      return false;
    end if;
    
    return p_v2 = p_v1.get_number();
  end;
  
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function equals(p_v1 pljson_value, p_v2 binary_double) return boolean as
  begin
    if (p_v2 is null) then
      return p_v1.is_null;
    end if;

    if (not p_v1.is_number) then
      return false;
    end if;

    return p_v2 = p_v1.get_double();
  end;
  
  function equals(p_v1 pljson_value, p_v2 boolean) return boolean as
  begin
    if (p_v2 is null) then
      return p_v1.is_null;
    end if;
    
    if (not p_v1.is_bool) then
      return false;
    end if;
    
    return p_v2 = p_v1.get_bool();
  end;
  
  function equals(p_v1 pljson_value, p_v2 varchar2) return boolean as
  begin
    if (p_v2 is null) then
      return (p_v1.is_null or p_v1.get_string() is null);
    end if;
    
    if (not p_v1.is_string) then
      return false;
    end if;
    
    return p_v2 = p_v1.get_string();
  end;
  
  function equals(p_v1 pljson_value, p_v2 clob) return boolean as
    my_clob clob;
    res boolean;
  begin
    if (p_v2 is null) then
      return p_v1.is_null;
    end if;
    
    if (not p_v1.is_string) then
      return false;
    end if;
    
--    
--    my_clob := empty_clob();
--    dbms_lob.createtemporary(my_clob, true);
--    p_v1.get_string(my_clob);
--    
    my_clob := p_v1.get_clob();
    res := dbms_lob.compare(p_v2, my_clob) = 0;
--    dbms_lob.freetemporary(my_clob);
    return res;
  end;
  
  function equals(p_v1 pljson_value, p_v2 pljson_value, exact boolean) return boolean as
  begin
    if (p_v2 is null or p_v2.is_null) then
      return (p_v1 is null or p_v1.is_null);
    end if;
    
    if (p_v2.is_number) then return equals(p_v1, p_v2.get_number); end if;
    if (p_v2.is_bool) then return equals(p_v1, p_v2.get_bool); end if;
    if (p_v2.is_object) then return equals(p_v1, pljson(p_v2), exact); end if;
    if (p_v2.is_array) then return equals(p_v1, pljson_list(p_v2), exact); end if;
    if (p_v2.is_string) then
      if (p_v2.extended_str is null) then
        return equals(p_v1, p_v2.get_string);
      else
        declare
          my_clob clob; res boolean;
        begin
--          
--          my_clob := empty_clob();
--          dbms_lob.createtemporary(my_clob, true);
--          p_v2.get_string(my_clob);
--          
          my_clob := p_v2.get_clob();
          res := equals(p_v1, my_clob);
--          dbms_lob.freetemporary(my_clob);
          return res;
        end;
      end if;
    end if;
    
    return false; --should never happen
  end;
  
  function equals(p_v1 pljson_value, p_v2 pljson_list, exact boolean) return boolean as
    cmp pljson_list;
    res boolean := true;
  begin
--  p_v1.print(false);
--  p_v2.print(false);
--  dbms_output.put_line('labc1'||case when exact then 'X' else 'U' end);
    
    if (p_v2 is null) then
      return p_v1.is_null;
    end if;
    
    if (not p_v1.is_array) then
      return false;
    end if;
    
--  dbms_output.put_line('labc2'||case when exact then 'X' else 'U' end);
    
    cmp := pljson_list(p_v1);
    if (cmp.count != p_v2.count and exact) then return false; end if;
    
--  dbms_output.put_line('labc3'||case when exact then 'X' else 'U' end);
    
    if (exact) then
      for i in 1 .. cmp.count loop
        res := equals(cmp.get(i), p_v2.get(i), exact);
        if (not res) then return res; end if;
      end loop;
    else
--  dbms_output.put_line('labc4'||case when exact then 'X' else 'U' end);
      if (p_v2.count > cmp.count) then return false; end if;
--  dbms_output.put_line('labc5'||case when exact then 'X' else 'U' end);
      
      --match sublist here!
      for x in 0 .. (cmp.count-p_v2.count) loop
--  dbms_output.put_line('labc7'||x);
        
        for i in 1 .. p_v2.count loop
          res := equals(cmp.get(x+i), p_v2.get(i), exact);
          if (not res) then
            goto next_index;
          end if;
        end loop;
        return true;
        
        <<next_index>>
        null;
      end loop;
    
--  dbms_output.put_line('labc7'||case when exact then 'X' else 'U' end);
    
    return false; --no match
    
    end if;
    
    return res;
  end;
  
  function equals(p_v1 pljson_value, p_v2 pljson, exact boolean) return boolean as
    cmp pljson;
    res boolean := true;
  begin
--  p_v1.print(false);
--  p_v2.print(false);
--  dbms_output.put_line('abc1');
    
    if (p_v2 is null) then
      return p_v1.is_null;
    end if;
    
    if (not p_v1.is_object) then
      return false;
    end if;
    
    cmp := pljson(p_v1);
    
--  dbms_output.put_line('abc2');
    
    if (cmp.count != p_v2.count and exact) then return false; end if;
    
--  dbms_output.put_line('abc3');
    declare
      k1 pljson_list := p_v2.get_keys();
      key_index number;
    begin
      for i in 1 .. k1.count loop
        key_index := cmp.index_of(k1.get(i).get_string());
        if (key_index = -1) then return false; end if;
        if (exact) then
          if (not equals(p_v2.get(i), cmp.get(key_index), true)) then return false; end if;
        else
          --non exact
          declare
            v1 pljson_value := cmp.get(key_index);
            v2 pljson_value := p_v2.get(i);
          begin
--  dbms_output.put_line('abc3 1/2');
--            v1.print(false);
--            v2.print(false);
            
            if (v1.is_object and v2.is_object) then
              if (not equals(v1, v2, false)) then return false; end if;
            elsif (v1.is_array and v2.is_array) then
              if (not equals(v1, v2, false)) then return false; end if;
            else
              if (not equals(v1, v2, true)) then return false; end if;
            end if;
          end;
        
        end if;
      end loop;
    end;
    
--  dbms_output.put_line('abc4');
    
    return true;
  end;
  
  function equals(p_v1 pljson, p_v2 pljson, exact boolean) return boolean as
  begin
    return equals(p_v1.to_json_value, p_v2, exact);
  end;
  
  function equals(p_v1 pljson_list, p_v2 pljson_list, exact boolean) return boolean as
  begin
    return equals(p_v1.to_json_value, p_v2, exact);
  end;
  
  --contain
  function contains(p_v1 pljson, p_v2 pljson_value, exact boolean) return boolean as
    v_values pljson_list;
  begin
    if (equals(p_v1.to_json_value, p_v2, exact)) then return true; end if;
    
    v_values := p_v1.get_values();
    
    for i in 1 .. v_values.count loop
      declare
        v_val pljson_value := v_values.get(i);
      begin
        if (v_val.is_object) then
          if (contains(pljson(v_val), p_v2, exact)) then return true; end if;
        end if;
        if (v_val.is_array) then
          if (contains(pljson_list(v_val), p_v2, exact)) then return true; end if;
        end if;
        
        if (equals(v_val, p_v2, exact)) then return true; end if;
      end;
    
    end loop;
     
    return false;
  end;
  
  function contains(p_v1 pljson_list, p_v2 pljson_value, exact boolean) return boolean as
  begin
    if (equals(p_v1.to_json_value, p_v2, exact)) then return true; end if;
    
    for i in 1 .. p_v1.count loop
      declare
        v_val pljson_value := p_v1.get(i);
      begin
        if (v_val.is_object) then
          if (contains(pljson(v_val), p_v2, exact)) then return true; end if;
        end if;
        if (v_val.is_array) then
          if (contains(pljson_list(v_val), p_v2, exact)) then return true; end if;
        end if;
        
        if (equals(v_val, p_v2, exact)) then return true; end if;
      end;
    
    end loop;
    
    return false;
  end;
  
  function contains(p_v1 pljson, p_v2 pljson, exact boolean ) return boolean as
  begin return contains(p_v1, p_v2.to_json_value, exact); end;
  function contains(p_v1 pljson, p_v2 pljson_list, exact boolean ) return boolean as
  begin return contains(p_v1, p_v2.to_json_value, exact); end;
  function contains(p_v1 pljson, p_v2 number, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function contains(p_v1 pljson, p_v2 binary_double, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
  function contains(p_v1 pljson, p_v2 varchar2, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
  function contains(p_v1 pljson, p_v2 boolean, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
  function contains(p_v1 pljson, p_v2 clob, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
  
  function contains(p_v1 pljson_list, p_v2 pljson, exact boolean ) return boolean as begin
  return contains(p_v1, p_v2.to_json_value, exact); end;
  function contains(p_v1 pljson_list, p_v2 pljson_list, exact boolean ) return boolean as begin
  return contains(p_v1, p_v2.to_json_value, exact); end;
  function contains(p_v1 pljson_list, p_v2 number, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
--   E.I.Sarmas (github.com/dsnz)   2016-12-01   support for binary_double numbers 
  function contains(p_v1 pljson_list, p_v2 binary_double, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
  function contains(p_v1 pljson_list, p_v2 varchar2, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
  function contains(p_v1 pljson_list, p_v2 boolean, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;
  function contains(p_v1 pljson_list, p_v2 clob, exact boolean ) return boolean as begin
  return contains(p_v1, pljson_value(p_v2), exact); end;


end pljson_helper;
/


--lshuo注释
--1.0存在会出错
begin execute immediate 'drop type pljson_vtab force'; exception when others then null; end;
/
create or replace type pljson_vtab as table of pljson_varray;
/

create or replace type pljson_table_impl as object (

  str clob, -- varchar2(32767),
--  
--    for 'nested' mode paths must use the [*] path operator
--  
  column_paths pljson_varray,
  column_names pljson_varray,
  table_mode varchar2(20),
  
--  
--    'cartesian' mode uses only
--    data_tab, row_ind
--  
  data_tab pljson_vtab,
--  
--    'nested' mode uses only
--    row_ind, row_count, nested_path
--    column_nested_index
--    last_nested_index
--    
--    for row_ind, row_count, nested_path
--    each entry corresponds to a [*] in the full path of the last column
--    and there will be the same or fewer entries than columns
--    1st nested path corresponds to whole array as '[*]'
--    or to root object as '' or to array within root object as 'key1.key2...array[*]'
--    
--    column_nested_index maps column index to nested_... index
--  
  row_ind pljson_narray,
  row_count pljson_narray,
--  
--    nested_path_full = full path, up to and including last [*], but not dot notation to key
--    nested_path_ext = extension to previous nested path
--    column_path_part = extension to nested_path_full, the dot notation to key after last [*]
--    column_path = nested_path_full || column_path_part
--    
--    start_column = start column where nested path appears first
--    nested_path_literal = nested_path_full with * replaced with literal integers, for fetching
--    
--    column_path = a[*].b.c[*].e
--    nested_path_full = a[*].b.c[*]
--    nested_path_ext = .b.c[*]
--    column_path_part = .e
--  
  nested_path_full pljson_varray,
  nested_path_ext pljson_varray,
  start_column pljson_narray,
  nested_path_literal pljson_varray,
  
  column_nested_index pljson_narray,
  column_path_part pljson_varray,
  column_val pljson_varray,
  
--   if the root of the document is array, the size of the array 
  root_array_size number,
  
--   the parsed json_obj 
  json_obj pljson,
  
  ret_type anytype,
  
  static function ODCITableDescribe(
    rtype out anytype,
    json_str clob, column_paths pljson_varray, column_names pljson_varray := null,
    table_mode varchar2 := 'cartesian'
  ) return number,
  
  static function ODCITablePrepare(
    sctx out pljson_table_impl,
    ti in sys.ODCITabFuncInfo,
    json_str clob, column_paths pljson_varray, column_names pljson_varray := null,
    table_mode varchar2 := 'cartesian'
  ) return number,
  
  static function ODCITableStart(
    sctx in out pljson_table_impl,
    json_str clob, column_paths pljson_varray, column_names pljson_varray := null,
    table_mode varchar2 := 'cartesian'
  ) return number,
  
  member function ODCITableFetch(
    self in out pljson_table_impl, nrows in number, outset out anydataset
  ) return number,
  
  member function ODCITableClose(self in pljson_table_impl) return number,
  
  static function json_table(
    json_str clob, column_paths pljson_varray, column_names pljson_varray := null,
    table_mode varchar2 := 'cartesian'
  ) return anydataset
  pipelined using pljson_table_impl
);
/
--lshuo注释
--增加Or replace,否则1.0存在会报错
create  or replace synonym pljson_table for pljson_table_impl;
create or replace type body pljson_table_impl as

  static function ODCITableDescribe(
    rtype out anytype,
    json_str clob, column_paths pljson_varray, column_names pljson_varray := null,
    table_mode varchar2 := 'cartesian'
  ) return number is
    atyp anytype;
  begin
    --dbms_output.put_line('>>Describe');
    
    anytype.begincreate(dbms_types.typecode_object, atyp);
    if column_names is null then
      for i in column_paths.FIRST .. column_paths.LAST loop
        atyp.addattr('JSON_' || ltrim(to_char(i)), dbms_types.typecode_varchar2, null, null, 32767, null, null);
      end loop;
    else
      for i in column_names.FIRST .. column_names.LAST loop
        atyp.addattr(upper(column_names(i)), dbms_types.typecode_varchar2, null, null, 32767, null, null);
      end loop;
    end if;
    atyp.endcreate;
    
    anytype.begincreate(dbms_types.typecode_table, rtype);
    rtype.SetInfo(null, null, null, null, null, atyp, dbms_types.typecode_object, 0);
    rtype.endcreate();
    
    --dbms_output.put_line('>>Describe end');
    return odciconst.success;
  exception
    when others then
      return odciconst.error;
  end;
  
  static function ODCITablePrepare(
    sctx out pljson_table_impl,
    ti in sys.ODCITabFuncInfo,
    json_str clob, column_paths pljson_varray, column_names pljson_varray := null,
    table_mode varchar2 := 'cartesian'
  ) return number is
    elem_typ sys.anytype;
    prec  pls_integer;
    scale pls_integer;
    len   pls_integer;
    csid  pls_integer;
    csfrm pls_integer;
    tc    pls_integer;
    aname varchar2(30);
  begin
    --dbms_output.put_line('>>Prepare');
    
    tc := ti.RetType.GetAttrElemInfo(1, prec, scale, len, csid, csfrm, elem_typ, aname);
    sctx := pljson_table_impl(
      json_str, column_paths, column_names,
      table_mode,
      pljson_vtab(), pljson_narray(), pljson_narray(),
      pljson_varray(), pljson_varray(),  pljson_narray(), pljson_varray(),
      pljson_narray(), pljson_varray(), pljson_varray(),
      0,
      pljson(),
      elem_typ
    );
    return odciconst.success;
  end;
  
  -- E.I.Sarmas (github.com/dsnz)   2017-09-23   NEW support for nested/cartesian table generation
  static function ODCITableStart(
    sctx in out pljson_table_impl,
    json_str clob, column_paths pljson_varray, column_names pljson_varray := null,
    table_mode varchar2 := 'cartesian'
  ) return number is
    json_obj pljson;
    json_val pljson_value;
    buf varchar2(32767);
    --data_tab pljson_vtab := pljson_vtab();
    json_arr pljson_list;
    json_elem pljson_value;
    value_array pljson_varray := pljson_varray();
    
    -- E.I.Sarmas (github.com/dsnz)   2017-09-23   NEW support for array as root json data
    root_val pljson_value;
    root_list pljson_list;
    root_array_size number := 0;
--     for nested mode 
    last_nested_path_full varchar2(32767);
    column_path varchar(32767);
    array_pos number;
    nested_path_prefix varchar2(32767);
    nested_path_ext varchar2(32767);
    column_path_part varchar2(32767);
--     a starts with b 
    function starts_with(a in varchar2, b in varchar2) return boolean is
    begin
      if b is null then
        return True;
      end if;
      if substr(a, 1, length(b)) = b then
        return True;
      end if;
      return False;
    end;
  begin
    -- json_obj := pljson(json_str);
    root_val := pljson_parser.parse_any(json_str);
    if root_val.typeval = 2 then
      root_list := pljson_list(root_val);
      root_array_size := root_list.count;
      json_obj := pljson(root_list);
    else
      -- implicit root of size 1
      root_array_size := 1;
      json_obj := pljson(root_val);
    end if;
  
    sctx.str := json_str;
    sctx.column_paths := column_paths;
    sctx.column_names := column_names;
    sctx.table_mode := table_mode;
    
    sctx.json_obj := json_obj;
    sctx.root_array_size := root_array_size;
    sctx.data_tab.delete;
    
    if table_mode = 'cartesian' then
      for i in column_paths.FIRST .. column_paths.LAST loop
        --dbms_output.put_line('path='||column_paths(i));
        json_val := pljson_ext.get_json_value(json_obj, column_paths(i));
        --dbms_output.put_line('type='||json_val.get_type());
        case json_val.typeval
          --when 1 then 'object';
          when 2 then
            json_arr := pljson_list(json_val);
            value_array.delete;
            for j in 1 .. json_arr.count loop
              json_elem := json_arr.get(j);
              case json_elem.typeval
                --when 1 then 'object';
                --when 2 then -- 'array';
                when 3 then
                  buf := json_elem.get_string();
                  value_array.extend(); value_array(value_array.LAST) := buf;
                when 4 then
                  buf := to_char(json_elem.get_number());
                  value_array.extend(); value_array(value_array.LAST) := buf;
                when 5 then
                  buf := case json_elem.get_bool() when true then 'true' when false then 'false' end;
                  value_array.extend(); value_array(value_array.LAST) := buf;
                when 6 then
                  buf := null;
                  value_array.extend(); value_array(value_array.LAST) := buf;
                else
                  -- if object is unknown or does not exist add new element of type null
                  buf := null;
                  --dbms_output.put_line('res[](unknown)='||buf);
                  sctx.data_tab.extend(); sctx.data_tab(sctx.data_tab.LAST) := pljson_varray(buf);
              end case;
            end loop;
            sctx.data_tab.extend(); sctx.data_tab(sctx.data_tab.LAST) := value_array;
          when 3 then 
            buf := json_val.get_string();
            sctx.data_tab.extend(); sctx.data_tab(sctx.data_tab.LAST) := pljson_varray(buf);
          when 4 then 
            buf := to_char(json_val.get_number());
            sctx.data_tab.extend(); sctx.data_tab(sctx.data_tab.LAST) := pljson_varray(buf);
          when 5 then 
            buf := case json_val.get_bool() when true then 'true' when false then 'false' end;
            sctx.data_tab.extend(); sctx.data_tab(sctx.data_tab.LAST) := pljson_varray(buf);
          when 6 then 
            buf := null;
            sctx.data_tab.extend(); sctx.data_tab(sctx.data_tab.LAST) := pljson_varray(buf);
          else
            -- if object is unknown or does not exist add new element of type null
            buf := null;
            --dbms_output.put_line('res(unknown)='||buf);
            sctx.data_tab.extend(); sctx.data_tab(sctx.data_tab.LAST) := pljson_varray(buf);
        end case;
      end loop;
      
      --dbms_output.put_line('initialize row indexes');
      sctx.row_ind.delete;
      --for i in data_tab.FIRST .. data_tab.LAST loop
      for i in column_paths.FIRST .. column_paths.LAST loop
        sctx.row_ind.extend();
        sctx.row_ind(sctx.row_ind.LAST) := 1;
      end loop;
    else
--       setup nested mode 
      sctx.nested_path_full.delete;
      sctx.nested_path_ext.delete;
      sctx.column_path_part.delete;
      sctx.column_nested_index.delete;
      for i in column_paths.FIRST .. column_paths.LAST loop
        --dbms_output.put_line(i || ', column_path = ' || column_paths(i));
        column_path := column_paths(i);
        array_pos := instr(column_path, '[*]', -1);
        if array_pos > 0 then
          nested_path_prefix := substr(column_path, 1, array_pos+2);
        else
          nested_path_prefix := '';
        end if;
        --dbms_output.put_line(i || ', nested_path_prefix = ' || nested_path_prefix);
        last_nested_path_full := '';
        if sctx.nested_path_full.LAST is not null then
          last_nested_path_full := sctx.nested_path_full(sctx.nested_path_full.LAST);
        end if;
        --dbms_output.put_line(i || ', last_nested_path_full = ' || last_nested_path_full);
        if not starts_with(nested_path_prefix, last_nested_path_full) then
          --dbms_output.put_line('column paths are not nested, column# ' || i);
          raise_application_error(-20120, 'column paths are not nested, column# ' || i);
        end if;
        if i = 1 or nested_path_prefix != last_nested_path_full
        or (nested_path_prefix is not null and last_nested_path_full is null) then
          nested_path_ext := substr(nested_path_prefix, nvl(length(last_nested_path_full), 0)+1);
          if instr(nested_path_ext, '[*]') != instr(nested_path_ext, '[*]', -1) then
            --dbms_output.put_line('column introduces more than one array, column# ' || i);
            raise_application_error(-20120, 'column introduces more than one array, column# ' || i);
          end if;
          sctx.nested_path_full.extend();
          sctx.nested_path_full(sctx.nested_path_full.LAST) := nested_path_prefix;
          --dbms_output.put_line(i || ', new nested_path_full = ' || nested_path_prefix);
          sctx.nested_path_ext.extend();
          sctx.nested_path_ext(sctx.nested_path_ext.LAST) := nested_path_ext;
          --dbms_output.put_line(i || ', new nested_path_ext = ' || nested_path_ext);
          sctx.start_column.extend();
          sctx.start_column(sctx.start_column.LAST) := i;
        end if;
        sctx.column_nested_index.extend();
        sctx.column_nested_index(sctx.column_nested_index.LAST) := sctx.nested_path_full.LAST;
        --dbms_output.put_line(i || ', column_nested_index = ' || sctx.nested_path_full.LAST);
        column_path_part := substr(column_path, nvl(length(nested_path_prefix), 0)+1);
        sctx.column_path_part.extend();
        sctx.column_path_part(sctx.column_path_part.LAST) := column_path_part;
        --dbms_output.put_line(i || ', column_path_part = ' || column_path_part);
      end loop;
      --dbms_output.put_line('initialize row indexes');
      sctx.row_ind.delete;
      sctx.row_count.delete;
      sctx.nested_path_literal.delete;
      sctx.column_val.delete;
      if sctx.nested_path_full.LAST is not null then
        for i in 1 .. sctx.nested_path_full.LAST loop
          sctx.row_ind.extend();
          sctx.row_ind(sctx.row_ind.LAST) := -1;
          sctx.row_count.extend();
          sctx.row_count(sctx.row_count.LAST) := -1;
          sctx.nested_path_literal.extend();
          sctx.nested_path_literal(sctx.nested_path_literal.LAST) := '';
        end loop;
      end if;
      for i in 1 .. sctx.column_paths.LAST loop
        sctx.column_val.extend();
        sctx.column_val(sctx.column_val.LAST) := '';
      end loop;
    end if;
    
    return odciconst.success;
  end;
  
  member function ODCITableFetch(
    self in out pljson_table_impl, nrows in number, outset out anydataset
  ) return number is
    --data_row pljson_varray := pljson_varray();
    --type index_array is table of number;
    --row_ind index_array := index_array();
    j number;
    num_rows number := 0;
    
    --json_obj pljson;
    json_val pljson_value;
    buf varchar2(32767);
    --data_tab pljson_vtab := pljson_vtab();
    json_arr pljson_list;
    json_elem pljson_value;
    value_array pljson_varray := pljson_varray();
    
--     nested mode 
    temp_path varchar(32767);
    start_index number;
    k number;
--    
--      k is nested path index and not column index
--      sets row_count()
--    
    procedure set_count(k number) is
      temp_path varchar(32767);
    begin
      if k = 1 then
        if nested_path_full(1) is null or nested_path_full(1) = '[*]' then
          row_count(1) := root_array_size;
          return;
        else
          temp_path := substr(nested_path_full(1), 1, length(nested_path_full(1)) - 3);
        end if;
      else
        temp_path := nested_path_literal(k - 1) || substr(nested_path_ext(k), 1, length(nested_path_ext(k)) - 3);
      end if;
      --dbms_output.put_line(k || ', set_count temp_path = ' || temp_path);
      json_val := pljson_ext.get_json_value(json_obj, temp_path);
      if json_val.typeval != 2 then
        raise_application_error(-20120, 'column introduces array with [*] but is not array in json, column# ' || k);
      end if;
      row_count(k) := pljson_list(json_val).count;
    end;
--    
--      k is nested path index and not column index
--      sets nested_path_literal() for row_ind(k)
--    
    procedure set_nested_path_literal(k number) is
      temp_path varchar(32767);
    begin
      if k = 1 then
        if nested_path_full(1) is null then
          return;
        end if;
        temp_path := substr(nested_path_full(1), 1, length(nested_path_full(1)) - 2);
      else
        temp_path := nested_path_literal(k - 1) || substr(nested_path_ext(k), 1, length(nested_path_ext(k)) - 2);
      end if;
      nested_path_literal(k) := temp_path || row_ind(k) || ']';
    end;
  begin
    
    if table_mode = 'cartesian' then
      outset := null;
      
      if row_ind(1) = 0 then
        return odciconst.success;
      end if;
      
      anydataset.begincreate(dbms_types.typecode_object, self.ret_type, outset);
      
--       iterative cartesian product algorithm 
      <<main_loop>>
      while True loop
        exit when num_rows = nrows or row_ind(1) = 0;
        --data_row.delete;
        outset.addinstance;
        outset.piecewise();
        for i in data_tab.FIRST .. data_tab.LAST loop
          --data_row.extend();
          --data_row(data_row.LAST) := data_tab(i)(row_ind(i));
          --dbms_output.put_line('json_'||ltrim(to_char(i)));
          --dbms_output.put_line('['||ltrim(to_char(row_ind(i)))||']');
          --dbms_output.put_line('='||data_tab(i)(row_ind(i)));
          outset.setvarchar2(data_tab(i)(row_ind(i)));
        end loop;
        --pipe row(data_row);
        num_rows := num_rows + 1;
        
        j := row_ind.COUNT;
        <<index_loop>>
        while True loop
          row_ind(j) := row_ind(j) + 1;
          if row_ind(j) <= data_tab(j).COUNT then
            exit index_loop;
          end if;
          row_ind(j) := 1;
          j := j - 1;
          if j < 1 then
            row_ind(1) := 0; -- hack to indicate end of all fetches
            exit main_loop;
          end if;
        end loop index_loop;
      end loop main_loop;
      
      outset.endcreate;
    else
--       fetch nested mode 
      outset := null;
      
      anydataset.begincreate(dbms_types.typecode_object, self.ret_type, outset);
      
      <<main_loop_nested>>
      while True loop
--         find starting column 
--        
--          in first run, loop will not assign value to start_index, so start_index := 0
--          in last run after all rows produced, the same will happen and start_index := 0
--          but the last run will have row_count(1) >= 0
--        
        start_index := 0;
        for i in REVERSE row_ind.FIRST .. row_ind.LAST loop
          if row_ind(i) < row_count(i) then
            start_index := start_column(i);
            exit;
          end if;
        end loop;
        if start_index = 0 then
          if num_rows = nrows or row_count(1) >= 0 then
            exit main_loop_nested;
          else
            start_index := 1;
          end if;
        end if;
        
--         fetch rows 
        <<row_loop_nested>>
        for i in start_index .. column_paths.LAST loop
          k := column_nested_index(i);
--           new nested path 
          if start_column(k) = i then
--             new count 
            if row_ind(k) = row_count(k) then
              set_count(k);
              row_ind(k) := 0;
              --dbms_output.put_line(i || ', new nested count = ' || row_count(k));
            end if;
--             advance row_ind 
            row_ind(k) := row_ind(k) + 1;
            set_nested_path_literal(k);
          end if;
          temp_path := nested_path_literal(k) || column_path_part(i);
          json_val := pljson_ext.get_json_value(json_obj, temp_path);
          case json_val.typeval
            when 3 then
              buf := json_val.get_string();
              column_val(i) := buf;
            when 4 then
              buf := to_char(json_val.get_number());
              column_val(i) := buf;
            when 5 then
              buf := case json_val.get_bool() when true then 'true' when false then 'false' end;
              column_val(i) := buf;
            when 6 then
              buf := null;
              column_val(i) := buf;
            else
              -- if object is unknown or does not exist add new element of type null
              buf := null;
              column_val(i) := buf;
          end case;
          if i = column_paths.LAST then
            outset.addinstance;
            outset.piecewise();
            for j in column_val.FIRST .. column_val.LAST loop
              outset.setvarchar2(column_val(j));
            end loop;
            num_rows := num_rows + 1;
          end if;
        end loop row_loop_nested;
      end loop main_loop_nested;
      
      outset.endcreate;
    end if;
    
    return odciconst.success;
  end;
  
  member function ODCITableClose(self in pljson_table_impl) return number is
  begin
    return odciconst.success;
  end;

end;
/

grant execute on pljson_element to public;
create or replace public synonym pljson_element for pljson_element;

grant execute on pljson_narray to public;
create or replace public synonym pljson_narray for pljson_narray;

grant execute on pljson_vtab to public;
create or replace public synonym pljson_vtab for pljson_vtab;

grant execute on pljson_varray to public;
create or replace public synonym pljson_varray for pljson_varray;

grant execute on pljson to public;
create or replace public synonym pljson for pljson;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json for pljson;

grant execute on pljson_list to public;
create or replace public synonym pljson_list for pljson_list;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_list for pljson_list;

grant execute on pljson_value_array to public;
create or replace public synonym pljson_value_array for pljson_value_array;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_value_array for pljson_value_array;

grant execute on pljson_value to public;
create or replace public synonym pljson_value for pljson_value;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_value for pljson_value;

grant execute on pljson_table_impl to public;
create or replace public synonym pljson_table_impl for pljson_table_impl;
create or replace public synonym pljson_table for pljson_table_impl;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_table for pljson_table;

--packages
grant execute on pljson_parser to public;
create or replace public synonym pljson_parser for pljson_parser;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_parser for pljson_parser;

grant execute on pljson_printer to public;
create or replace public synonym pljson_printer for pljson_printer;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_printer for pljson_printer;

grant execute on pljson_ext to public;
create or replace public synonym pljson_ext for pljson_ext;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_ext for pljson_ext;

grant execute on pljson_dyn to public;
create or replace public synonym pljson_dyn for pljson_dyn;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_dyn for pljson_dyn;

grant execute on pljson_ml to public;
create or replace public synonym pljson_ml for pljson_ml;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_ml for pljson_ml;

grant execute on pljson_xml to public;
create or replace public synonym pljson_xml for pljson_xml;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_xml for pljson_xml;

grant execute on pljson_util_pkg to public;
create or replace public synonym pljson_util_pkg for pljson_util_pkg;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_util_pkg for pljson_util_pkg;

grant execute on pljson_helper to public;
create or replace public synonym pljson_helper for pljson_helper;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_helper for pljson_helper;

grant execute on pljson_ac to public;
create or replace public synonym pljson_ac for pljson_ac;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create or replace public synonym json_ac for pljson_ac;
--lshuo注释
--不需要单元测试，因此注销
--grant execute on pljson_ut to public;
--create or replace public synonym pljson_ut for pljson_ut;
--grant all on pljson_testsuite to public;
--create or replace public synonym pljson_testsuite for pljson_testsuite;
--lshuo注释
--保证1.0能够和3.0共存，此处不创建同义词
--create synonym json_parser for pljson_parser;
--create synonym json_printer for pljson_printer;
--create synonym json_ext for pljson_ext;
--create synonym json_dyn for pljson_dyn;
--create synonym json_ml for pljson_ml;
--create synonym json_xml for pljson_xml;
--create synonym json_util_pkg for pljson_util_pkg;
--create synonym json_helper for pljson_helper;
--create synonym json_ac for pljson_ac;
--create synonym json for pljson;
--create synonym json_list for pljson_list;
--create synonym json_value_array for pljson_value_array;
--create synonym json_value for pljson_value;
--create synonym json_table for pljson_table;

--lshuo注释
--当存在1.0时，会导致JSON_TABLE同义词失效
Declare
  n_Have1version Number(3);
Begin
  Select Count(1)
  Into n_Have1version
  From All_Objects A
  Where a.Object_Name Like 'JSON_LIST' And a.Object_Type = 'TYPE' And a.Owner = User;
  If n_Have1version > 0 Then
    Begin Execute Immediate 'create or replace synonym json_table for pljson_table';Exception When Others Then Null;End;
    Begin Execute Immediate 'create or replace public synonym json_table for pljson_table';Exception When Others Then Null;End;
  End If;
End;
/

