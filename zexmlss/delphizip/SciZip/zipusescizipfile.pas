//****************************************************************
// Routines for pack/unpack to/from zip using SciZipFile
// Author:  Daniel Schnabel
// Last update: 2015-11-05
//----------------------------------------------------------------

unit zipusescizipfile;

interface

uses
  SysUtils,
  SciZipFile;

function ZEZipUnpackSciZipFile(ZipName: string; PathName: string): boolean;
function ZEZipPackSciZipFile(ZipName: string; PathName: string): boolean;

implementation

uses
  Classes;

//Unpack zip ZipName to folder PathName
//INPUT
//      ZipName: string   - full name of zip file
//      PathName: string  - path
//RETURN
//      boolean - true - ok
function ZEZipUnpackSciZipFile(ZipName: string; PathName: string): boolean;
var
  zip: TZipFile;
  i: integer;
  fn: string;
  sl: TStringList;
begin
  Result := false;

  if (not FileExists(ZipName)) then
   begin
     Exit;
   end;

  PathName := IncludeTrailingPathDelimiter(PathName);

  if (not DirectoryExists(PathName)) then
   begin
     Exit;
   end;

  zip := TZipFile.Create;
  try
    zip.LoadFromFile(ZipName);

    sl := TStringList.Create;
    try
      for i := 0 to zip.Count - 1 do
       begin
         fn := StringReplace(String(zip.Name[i]), '/', PathDelim, [rfReplaceAll]);
         ForceDirectories(ExtractFileDir(IncludeTrailingPathDelimiter(PathName) + fn));

         sl.Text := String(zip.Data[i]);
         sl.SaveToFile(IncludeTrailingPathDelimiter(PathName) + fn);
       end;
    finally
      sl.Free;
    end;

    Result := true;
  finally
    zip.Free;
  end;
end;

// Pack all the contents of a folder in PathName ZipName archive (used SciZipFile)
// INPUT
// ZipName: string - the name for the backup
// PathName: string - the folder that needs to be archived
// RETURN
// Boolean - true - file successfully created
function ZEZipPackSciZipFile(ZipName: string; PathName: string): boolean;

  procedure GetAllFilesInDirectory(Dir: string; FileList: TStringList);
  var
    sr: TSearchRec;
    thisdir: string;
  begin
    thisdir := IncludeTrailingPathDelimiter(Dir);

    if FindFirst(thisdir + '*.*', faAnyFile, sr) = 0 then
     begin
       try
         repeat
           if (sr.Attr and faDirectory) = faDirectory then
            begin
              if (sr.Name <> '..') and (sr.Name <> '.') then
               begin
                 GetAllFilesInDirectory(thisdir + sr.Name, FileList);
               end;
            end
           else
            begin
              FileList.Add(thisdir + sr.Name);
            end;
         until FindNext(sr) <> 0;
       finally
         FindClose(sr);
       end;
     end;
  end;

var
  zip: TZipFile;
  i: integer;
  s: string;
  sl: TStringList;
  stream: TFileStream;
  filetoadd: string;
  buffer: AnsiString;
begin
  Result := false;

  PathName := IncludeTrailingPathDelimiter(PathName);
  s := ExtractFilePath(ZipName);

  if not ForceDirectories(s) then
   begin
     Exit;
   end;

  sl := TStringList.Create;
  try
    GetAllFilesInDirectory(PathName, sl);

    zip := TZipFile.Create;
    try
      for i := 0 to sl.Count - 1 do
       begin
         filetoadd := sl[i];
         Delete(filetoadd, 1, Length(PathName));
         zip.AddFile(AnsiString(filetoadd));

         stream := TFileStream.Create(sl[i], fmOpenRead);
         try
           SetLength(buffer, stream.Size);
           stream.ReadBuffer(buffer[1], stream.Size);
           zip.Data[zip.Count - 1] := buffer;
         finally
           stream.Free;
         end;
       end;

      zip.SaveToFile(ZipName);

      Result := true;
    finally
      zip.Free;
    end;
  finally
    sl.Free;
  end;
end;

end.
