//****************************************************************
// Routines for pack/unpack to/from zip using JEDI Code Library (via 7z)
// http://sourceforge.net/projects/jcl/ (under MPL 1.1).
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib ?
// Last update: 2014.11.22
//----------------------------------------------------------------
{
 Copyright (C) 2014 Ruslan Neborak

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

unit zipusejcl7z;

interface

uses
  SysUtils,
  JclCompression;

function ZEZipUnpackJCL7z(ZipName: string; PathName: string): boolean;
function ZEZipPackJCL7z(ZipName: string; PathName: string): boolean;

implementation

//Unpack zip ZipName to folder PathName
//INPUT
//      ZipName: string   - full name of zip file
//      PathName: string  - path
//RETURN
//      boolean - true - ok
function ZEZipUnpackJCL7z(ZipName: string; PathName: string): boolean;
var
  t: integer;
  _un: TJclZipDecompressArchive;
  
begin
  Result := false;

  if (not FileExists(ZipName)) then
    exit;

  t := Length(PathName);
  if (t > 0) then
    if (PathName[t] <> PathDelim) then
      PathName := PathName + PathDelim;

  if (not DirectoryExists(PathName)) then
    exit;

  _un := nil;
  try
    _un := TJclZipDecompressArchive.Create(ZipName);
    _un.ListFiles();
    _un.ExtractAll(PathName, true);
  finally
    if (Assigned(_un)) then
      FreeAndNil(_un);
  end;

end; //ZEZipUnpackJCL7z

function ZEZipPackJCL7z(ZipName: string; PathName: string): boolean;
var
  _packer: TJclZipCompressArchive;
  i: integer;
  s: string;
  _pathlen: integer;

  procedure AddFolder(FolderName: string);
  var
    sr: TSearchRec;

  begin
    result := true;
    if (FindFirst(FolderName + '*.*', faAnyFile, sr) = 0) then
    try
      repeat
        if ((sr.Attr and faDirectory) = faDirectory) then
        begin
          if (sr.Name <> '..') and (sr.Name <> '.') then
            AddFolder(FolderName + sr.Name + PathDelim);
        end
        else
        begin
          s := FolderName + sr.Name;
          delete(s, 1, _pathlen);
          _packer.AddFile(s, FolderName + sr.Name);
        end;
      until FindNext(sr) <> 0;
    finally
      FindClose(sr);
    end;
  end; //AddFolder

begin
  Result := false;

  i := length(PathName);
  if (i > 0) then
    if (PathName[i] <> PathDelim) then
      PathName := PathName + PathDelim;

  _pathlen := Length(PathName);    

  _packer := nil;
  try
    s := ExtractFilePath(ZipName);
    if (not ForceDirectories(s)) then
      exit;

    _packer := TJclZipCompressArchive.Create(ZipName);
    //I do not understand how works AddDirectory...
    //_packer.AddDirectory('', PathName, true, true);
    AddFolder(PathName);
    _packer.Compress();

    Result := true
  finally
    if (Assigned(_packer)) then
      FreeAndNil(_packer);
  end;
end; //ZEZipPackJCL7z

end.
