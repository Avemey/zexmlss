//****************************************************************
// Routines for pack/unpack to/from zip using Kazip
// https://raw.githubusercontent.com/JoseJimeniz/KaZip/master/KaZip.pas
// (original Kazip: http://kadao.dir.bg/download/kazip/)
// Author:  Anonym
// e-mail:  ???
// URL:     ???
// License: zlib ?
// Last update: 2015.03.29
//----------------------------------------------------------------
{
 Copyright (C) 2015 Anonym

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
unit zipusekazip;

interface

uses
  SysUtils, Kazip;

function ZEZipUnpackKazip(ZipName: string; PathName: string): boolean;
function ZEZipPackKazip(ZipName: string; PathName: string): boolean;

implementation

//–аспаковывет zip архив ZipName в путь PathName (используетс€ Kazip)
//INPUT
//      ZipName: string   - им€ архива
//      PathName: string  - папка, куда нужно распаковать архив. ѕуть должен существовать.
//RETURN
//      boolean - true - архив распакован успешно
function ZEZipUnpackKaZip(ZipName: string; PathName: string): boolean;
var
  t: integer;
  zip: TKaZip;
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

  zip := TKaZip.Create(nil);
  try
    zip.Open(ZipName);
    zip.OverwriteAction := oaOverwriteAll;
    zip.ExtractAll(PathName);
    zip.Close;
    result := true;
  finally
    zip.Free;
  end;
end; //ZEZipUnpackKaZip

//«апаковывает всЄ содержимое папки PathName в архив ZipName (используетс€ KaZip)
//ѕуть PathName должен существовать.
//INPUT
//      ZipName: string   - им€ создаваемого архива
//      PathName: string  - что нужно заархивировать
//RETURN
//      boolean - true - архив успешно создан
function ZEZipPackKaZip(ZipName: string; PathName: string): boolean;
var
  zip: TKaZip;
  i: integer;
  s: string;

begin
  result := false;

  i := length(PathName);
  if (i > 0) then
    if (PathName[i] <> PathDelim) then
      PathName := PathName + PathDelim;

  s := ExtractFilePath(ZipName);
  if (not ForceDirectories(s)) then
    Exit;

  zip := TKaZip.Create(nil);
  try
    zip.CreateZip(ZipName);
    zip.Open(ZipName);
    result := zip.AddFolder(PathName, PathName, '*.*', True);
    zip.Close;
  finally
    zip.Free;
  end;
end; //ZEZipPackKaZip

end.
