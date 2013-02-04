//****************************************************************
// Routines for pack/unpack to/from zip using Synopse
// http://synopse.info (licensed under a MPL/GPL/LGPL tri-license).
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib ?
// Last update: 2013.01.16
//----------------------------------------------------------------
{
 Copyright (C) 2013 Ruslan Neborak

  This software is provided 'as-is', without any express or implied
 warranty. In no event will the authors be held liable for any damages
 arising from the use of this software.

 Permission is granted to anyone to use this software for any purpose,
 including commercial applications, and to alter it and redistribute it
 freely, subject to the following restrictions:

    1. The origin of this software must not be misrepresented; you must not
    claim that you wrote the original software. If you use this software
    in a product, an acknowledgment in the product documentation would be
    appreciated but is not required.

    2. Altered source versions must be plainly marked as such, and must not be
    misrepresented as being the original software.

    3. This notice may not be removed or altered from any source
    distribution.
}
unit zipusesyn;

interface

uses
  SysUtils, classes, SynZip, SynZipFiles;

function ZEZipPackSyn(ZipName: string; PathName: string): boolean;
function ZEZipUnpackSyn(ZipName: string; PathName: string): boolean;

implementation

//–аспаковывет zip архив ZipName в путь PathName (используетс€ SynZip)
//INPUT
//      ZipName: string   - им€ архива
//      PathName: string  - папка, куда нужно распаковать архив. ѕуть должен существовать.
//RETURN
//      boolean - true - архив распакован успешно
function ZEZipUnpackSyn(ZipName: string; PathName: string): boolean;
var
  zr: TZipReader;
  i: integer;
  stream: TStream;
  s: string;
  t: integer;
  _fname: string;

begin
  result := false;
  if (not FileExists(ZipName)) then
    exit;

  t := Length(PathName);
  if (t > 0) then
    if (PathName[t] <> PathDelim) then
      PathName := PathName + PathDelim;

  if (not DirectoryExists(PathName)) then
    exit;

  zr := nil;
  try
    zr := TZipReader.Create(ZipName);
    for i := 0 to zr.Count - 1 do
    begin
      stream := nil;
      _fname := zr.Entry[i].ZipName;

      s := ExtractFilePath(_fname);

      if (length(s) > 0) then
        ForceDirectories(PathName + s);

      if (not zr.Entry[i].Header.IsFolder) then
      try
        stream := TFileStream.Create(PathName + _fname, fmCreate);
        zr.GetData(i, stream);
      finally
        if (Assigned(stream)) then
          FreeAndNil(stream);
      end;
    end;

    result := true

  finally
    if (Assigned(zr)) then
      FreeAndNil(zr);
  end;
end; //ZEZipUnpackSyn

//«апаковывает всЄ содержимое папки PathName в архив ZipName (используетс€ SynZip)
//ѕуть PathName должен существовать.
//INPUT
//      ZipName: string   - им€ создаваемого архива
//      PathName: string  - что нужно заархивировать
//RETURN
//      boolean - true - архив успешно создан
function ZEZipPackSyn(ZipName: string; PathName: string): boolean;
type
  ttt = array of string;

var
  tz: TZipWriter;
  z: array of string;
  kol, i: integer;
  s: string;

  function GetFilesList(DirName: string; var FileList: ttt): integer;
  var
    _FilesCount: integer;
    _ArraySize: integer;

    procedure DoRecurce(const adir: string);
    var
      sr: TSearchRec;

    begin
      if FindFirst(DirName + adir + '*.*', faAnyFile, sr) = 0 then
      begin
        repeat
          if ((sr.Attr and faDirectory) = faDirectory) then
          begin
            if (sr.Name <> '..') and (sr.Name <> '.') then
              DoRecurce(adir + sr.Name + PathDelim);
          end else
          begin
            if (_FilesCount >= _ArraySize) then
            begin
              _ArraySize := _FilesCount + 1;
              SetLength(FileList, _ArraySize);
            end;
            FileList[_FilesCount] := adir + sr.Name;
            inc(_FilesCount);
          end;
        until FindNext(sr) <> 0;
        FindClose(sr);
      end;
    end; //DoRecurce

  begin
    _FilesCount := 0;
    _ArraySize := High(FileList) + 1;
    try
      DoRecurce('');
    finally
      result := _FilesCount;
    end;
  end; //GetFilesList

  function _replace(const st: string): string;
  var
    i: integer;

  begin
    result := '';
    for i := 1 to length(st) do
      if (st[i] = '/') then
        result := result + '\'
      else
        result := result + st[i];
  end; //_replace

begin
  tz := nil;
  result := false;
  i := length(PathName);
  if (i > 0) then
    if (PathName[i] <> PathDelim) then
      PathName := PathName + PathDelim;

  kol := GetFilesList(PathName, ttt(z));

  try
    s := ExtractFilePath(ZipName);
    if (not ForceDirectories(s)) then
      exit;

    tz := TZipWriter.Create(ZipName);

    for i := 0 to kol - 1 do
      tz.AddFile(PathName + z[i], _replace(z[i]), 6);

    result := true;
  finally
    if (Assigned(tz)) then
      FreeAndNil(tz);
    SetLength(z, 0);
    z := nil;
  end;
end; //ZEZipPackSyn

end.
