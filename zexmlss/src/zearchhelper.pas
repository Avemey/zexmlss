//****************************************************************
// Common routines for pack/unpack to/from zip.
// Author:  Ruslan V. Neborak
// e-mail:  avemey@tut.by
// URL:     http://avemey.com
// License: zlib
// Last update: 2013.01.12
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

unit zearchhelper;

interface

uses
  {$ifndef FPC}Windows,{$endif}
  Classes,
  SysUtils;

{$ifdef MSWINDOWS}
type
  TTmpFileStream = class(THandleStream)
  public
    constructor Create();
    destructor Destroy; override;
  end;
{$endif}

function ZEGetTempDir(): string;
function ZECreateUniqueTmpDir(const ADir: string; var retTmpDir: string): boolean;
function ZEDelTree(ADir: string): boolean;
// ����������� �����
function ZECopyFile(const ASource, ADest: TFileName): Boolean;

implementation

//�������� ���� � �����
function ZEGetTempDir(): string;
begin
  {$IFDEF FPC}
  result := GetTempDir();
  {$ELSE}
  result := GetEnvironmentVariable('TEMP');
  {$ENDIF}
end;

//������ ���������� ������ ���������� � ADir
//INPUT
//  const ADir: string    - ����, � ������� ����� ������� ����������
//  var retTmpDir: string - ������������ ���� � ��������� ����������
//RETURN
//        boolean - true - ���������� ������� �������
function ZECreateUniqueTmpDir(const ADir: string; var retTmpDir: string): boolean;
var
  l: integer;
  stime, s: string;
  _now: TDateTime;
  _y, _m, _d: word;
  _h, _min, _s, _ms: word;
  kol: integer;

begin
  result := false;
  l := length(ADir);
  retTmpDir := ADir;
  if (l > 1) then
    if (ADir[l] <> PathDelim) then
      retTmpDir := retTmpDir + PathDelim;

  _now := now();
  DecodeDate(_now, _y, _m, _d);
  DecodeTime(_now, _h, _min, _s, _ms);

  stime := '_tmp' + IntToStr(_y) + IntToStr(_m) + IntToStr(_d) +
           IntToStr(_h) + IntToStr(_min) + IntToStr(_s) + IntToStr(_ms) + '_';
  s := '';
  kol := 0;
  repeat
    s := s + IntToHex(random(16), 1);
    inc(kol);
    result := ForceDirectories(retTmpDir + stime + s);
  until (result) or (kol > 40);

  if (result) then
    retTmpDir := retTmpDir + stime + s + PathDelim;
end; //ZECreateUniqueTmpDir

//������� ���������� � ����������
//INPUT
//      ADir: string - ��� ��������� ����������
//RETURN
//      boolean - true - ���������� ������� �������
function ZEDelTree(ADir: string): boolean;

  function _DelTree(const AddDir: string): boolean;
  var
      sr: TSearchRec;

  begin
    result := true;
    if FindFirst({ADir + }AddDir + '*.*', faAnyFile, sr) = 0 then
    try
      repeat
        if ((sr.Attr and faDirectory) = faDirectory) then
        begin
          if (sr.Name <> '..') and (sr.Name <> '.') then
            _DelTree({ADir + }AddDir + sr.Name + PathDelim);
        end else
          result := result and DeleteFile({ADir +} AddDir + sr.Name);
      until FindNext(sr) <> 0;
    finally
      FindClose(sr);
      if (result) then
        result := result and RemoveDir({ADir +} AddDir);
    end;
  end;

begin
  result := _DelTree(ADir{''});
end;

// ����������� �����
function ZECopyFile(const ASource, ADest: TFileName): Boolean;
{$ifdef FPC}
var
  fs1, fs2: TFileStream;
{$endif}
begin
  {$ifndef FPC}
  Result := Windows.CopyFile(PChar(ASource), PChar(ADest), False);
  {$else}
  Result := False;
  if not FileExists(ASource) then Exit;

  if FileExists(ADest) then
    DeleteFile(ADest);

  fs1 := TFileStream.Create(ASource, fmOpenRead);
  try
    fs2 := TFileStream.Create(ADest, fmCreate);
    try
      fs2.CopyFrom(fs1, fs1.Size);
    finally
      fs2.Free();
    end;
  finally
    fs1.Free();
  end;
  {$endif}
end;

{ TTmpFileStream }

{$ifdef MSWINDOWS}
constructor TTmpFileStream.Create;
var
  TmpGUID: TGuid;
  FileName: string;
  tempFolder: array[0..MAX_PATH] of Char;
begin
  GetTempPath(MAX_PATH, @tempFolder);
  CreateGUID(TmpGuid);
  FileName := StrPas(tempFolder) + GUIDToString(TmpGuid);

  inherited Create(
    CreateFile(PChar(FileName), GENERIC_READ or GENERIC_WRITE, 0, nil, CREATE_ALWAYS,
    FILE_ATTRIBUTE_TEMPORARY or FILE_FLAG_DELETE_ON_CLOSE, 0)
  );
  if Handle < 0 then
    raise EFCreateError.CreateFmt('Cannot create file "%s". %s', [ExpandFileName(FileName), SysErrorMessage(GetLastError)]);
end;

destructor TTmpFileStream.Destroy;
begin
  if Handle >= 0 then FileClose(Handle);
  inherited Destroy;
end;
{$endif}

end.
