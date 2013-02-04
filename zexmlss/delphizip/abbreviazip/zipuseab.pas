//****************************************************************
// Routines for pack/unpack to/from zip using TurboPower Abbrevia
// http://sourceforge.net/projects/tpabbrevia/ (under MPL 1.1).
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
unit zipuseab;

interface

uses
  SysUtils, AbZipper, AbUnzper, AbArcTyp, AbUtils;

function ZEZipUnpackAb(ZipName: string; PathName: string): boolean;
function ZEZipPackAb(ZipName: string; PathName: string): boolean;

implementation

//–аспаковывет zip архив ZipName в путь PathName (используетс€ AbZipper)
//INPUT
//      ZipName: string   - им€ архива
//      PathName: string  - папка, куда нужно распаковать архив. ѕуть должен существовать.
//RETURN
//      boolean - true - архив распакован успешно
function ZEZipUnpackAb(ZipName: string; PathName: string): boolean;
var
  t: integer;
  ab: TAbUnZipper;

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

  ab := nil;
  try
    ab := TAbUnZipper.Create(nil);
    ab.FileName := ZipName;
    ab.Logging := false;
    ab.BaseDirectory := PathName;
    ab.ExtractOptions := [eoCreateDirs, eoRestorePath];
    ab.ExtractFiles('*.*');
    ab.CloseArchive();
    result := true;
  finally
    if (Assigned(ab)) then
      FreeAndNil(ab);
  end;
end; //ZEZipUnpackAb

//«апаковывает всЄ содержимое папки PathName в архив ZipName (используетс€ AbZipper)
//ѕуть PathName должен существовать.
//INPUT
//      ZipName: string   - им€ создаваемого архива
//      PathName: string  - что нужно заархивировать
//RETURN
//      boolean - true - архив успешно создан
function ZEZipPackAb(ZipName: string; PathName: string): boolean;
var
  ab: TAbZipper;
  i: integer;
  s: string;

begin
  result := false;
  ab := nil;

  i := length(PathName);
  if (i > 0) then
    if (PathName[i] <> PathDelim) then
      PathName := PathName + PathDelim;

  try
    s := ExtractFilePath(ZipName);
    if (not ForceDirectories(s)) then
      exit;

    ab := TAbZipper.Create(nil);
    ab.ArchiveType := atZip;
    ab.ForceType := true;
    ab.FileName := ZipName;
    ab.BaseDirectory := PathName;
    ab.StoreOptions := [soRecurse];
    ab.AddFiles('*', faAnyFile);
    ab.Save();
    ab.CloseArchive();
    result := true
  finally
    if (Assigned(ab)) then
      FreeAndNil(ab);
  end;
end; //ZEZipPackAb

end.
