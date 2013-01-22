//****************************************************************
// Routines for ansistrings for delphi >=2009
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2012.03.31
//----------------------------------------------------------------

unit duansistr;

interface

function DUAnsiPos(const Substr: ansistring; const S: ansistring): integer;

implementation

function DUAnsiPos(const Substr: ansistring; const S: ansistring): integer;
var
  i, j: integer;
  kol_str: integer;
  kol_sub: integer;
  b: boolean;

begin
  kol_str := Length(s);
  kol_sub := Length(substr);
  result := 0;
  if (kol_sub = 0) then
    exit;
  for i := 1 to kol_str - kol_sub + 1 do
    if (s[i] = Substr[1]) then
    begin
      b := true;
      for j := 2 to kol_sub do
        if (s[i + j - 1] <> Substr[j]) then
        begin
          b := false;
          break;
        end;
      if (b) then
      begin
        result := i;
        break;
      end;
    end;
end;

end.
