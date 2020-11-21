(*
 FPC Zipper imitation for Delphi
*)
unit zeZipper;

interface

{$ifdef FPC}
  {$error For Delphi only!}
{$endif}

{$i zexml.inc}

uses
  Classes, SysUtils
  {$ifdef XE2ZIP}, System.Zip{$endif}
  {$ifdef KAZIP}, KAZip{$endif}
  {$ifdef JCL7Z}, JclCompression{$endif}
  {$ifdef ABZIP}, AbZipper, AbUnzper, AbArcTyp, AbUtils{$endif}
  {$ifdef SCIZIP}, SciZipFile{$endif}
  ;

type
  TProgressEvent = procedure(Sender: TObject; const Pct: Double) of object;
  TOnEndOfFileEvent = procedure(Sender: TObject; const Ratio: Double) of object;
  TOnStartFileEvent = procedure(Sender: TObject; const AFileName: String) of object;

  EZipError = class(Exception);

  TCompressionLevel = (
    clNone,                     {Do not use compression, just copy data.}
    clFastest,                  {Use fast (but less) compression.}
    clDefault,                  {Use default compression}
    clMax                       {Use maximum compression}
  );

  { TZipFileEntry }

  TZipFileEntry = class(TCollectionItem)
  private
    FArchiveFileName: String; //Name of the file as it appears in the zip file list
    FAttributes: LongWord;
    FDateTime: TDateTime;
    FDiskFileName: String; {Name of the file on disk (i.e. uncompressed. Can be empty if based on a stream.);
                            uses local OS/filesystem directory separators}
    FOS: Byte;
    FSize: Int64;
    FStream: TStream;
    FCompressionLevel: TCompressionlevel;
    function GetArchiveFileName: String;
    procedure SetArchiveFileName(const AValue: String);
    procedure SetDiskFileName(const AValue: String);
  public
    constructor Create(ACollection: TCollection); override;
    function IsDirectory: Boolean;
    function IsLink: Boolean;
    procedure Assign(Source: TPersistent); override;
    property Stream: TStream read FStream write FStream;
  published
    property ArchiveFileName: String read GetArchiveFileName write SetArchiveFileName;
    property DiskFileName: String read FDiskFileName write SetDiskFileName;
    property Size: Int64 read FSize write FSize;
    property DateTime: TDateTime read FDateTime write FDateTime;
    property OS: Byte read FOS write FOS;
    property Attributes: LongWord read FAttributes write FAttributes;
    Property CompressionLevel: TCompressionlevel read FCompressionLevel write FCompressionLevel;
  end;

  { TZipFileEntries }

  TZipFileEntries = class(TCollection)
  private
    function GetZ(AIndex: Integer): TZipFileEntry;
    procedure SetZ(AIndex: Integer; const AValue: TZipFileEntry);
  public
    function AddFileEntry(const ADiskFileName: String): TZipFileEntry; overload;
    function AddFileEntry(const ADiskFileName, AArchiveFileName: String): TZipFileEntry; overload;
    function AddFileEntry(const AStream: TSTream; const AArchiveFileName: String): TZipFileEntry; overload;
    procedure AddFileEntries(const List: TStrings);
    property Entries[AIndex: Integer]: TZipFileEntry read GetZ write SetZ; default;
  end;

  { TFullZipFileEntry }

  TFullZipFileEntry = class(TZipFileEntry)
  private
    FBitFlags: Word;
    FCompressedSize: Int64;
    FCompressMethod: Word;
    FCRC32: LongWord;
  public
    property BitFlags: Word read FBitFlags;
    property CompressMethod: Word read FCompressMethod;
    property CompressedSize: Int64 read FCompressedSize;
    property CRC32: LongWord read FCRC32 write FCRC32;
  end;

  TOnCustomStreamEvent = procedure(Sender: TObject; var AStream: TStream; AItem: TFullZipFileEntry) of object;
  TCustomInputStreamEvent = procedure(Sender: TObject; var AStream: TStream) of object;

  { TFullZipFileEntries }

  TFullZipFileEntries = class(TZipFileEntries)
  private
    function GetFZ(AIndex: Integer): TFullZipFileEntry;
    procedure SetFZ(AIndex: Integer; const AValue: TFullZipFileEntry);
  public
    property FullEntries[AIndex: Integer]: TFullZipFileEntry read GetFZ write SetFZ; default;
  end;

  { TZipper }

  TZipper = class(TObject)
  private
    FEntries        : TZipFileEntries;
    FZipping        : Boolean;
    FBufSize        : LongWord;
    FFileName       : String;         { Name of resulting Zip file                 }
    FFileComment    : String;
    FFiles          : TStrings;
    FInMemSize      : Int64;
    FTmpZipFolder   : string;
    //FCompressor     : TObject;
    FOnPercent      : LongInt;
    FOnProgress     : TProgressEvent;
    FOnEndOfFile    : TOnEndOfFileEvent;
    FOnStartFile    : TOnStartFileEvent;
    function CheckEntries: Integer;
    procedure SetEntries(const AValue: TZipFileEntries);
    {$ifdef KAZIP}procedure BuildZipDirectoryKaZip();{$endif}
    {$ifdef JCL7Z}procedure BuildZipDirectoryJCL7Z();{$endif}
    {$ifdef ABZIP}procedure BuildZipDirectoryAbZip();{$endif}
    {$ifdef SCIZIP}procedure BuildZipDirectorySciZip();{$endif}
    {$ifdef XE2ZIP}procedure BuildZipDirectoryXE2Zip();{$endif}
  protected
    //Builds central directory based on local headers
    procedure BuildZipDirectory;
    procedure GetFileInfo();
    procedure SetBufSize(Value: LongWord);
    procedure SetFileName(Value: String);
  public
    constructor Create;
    destructor Destroy; override;
    procedure ZipAllFiles; virtual;
    // Saves zip to file and changes FileName
    procedure SaveToFile(AFileName: string);
    // Saves zip to stream
    procedure SaveToStream(AStream: TStream);
    // Zips specified files into a zip with name AFileName
    procedure ZipFiles(AFileName: String; FileList: TStrings); overload;
    procedure ZipFiles(FileList: TStrings); overload;
    // Zips specified entries into a zip with name AFileName
    procedure ZipFiles(AFileName: String; Entries: TZipFileEntries); overload;
    procedure ZipFiles(Entries: TZipFileEntries); overload;
    procedure Clear;
  public
    property BufferSize: LongWord read FBufSize write SetBufSize;
    property OnPercent: Integer read FOnPercent write FOnPercent;
    property OnProgress: TProgressEvent read FOnProgress write FOnProgress;
    property OnStartFile: TOnStartFileEvent read FOnStartFile write FOnStartFile;
    property OnEndFile: TOnEndOfFileEvent read FOnEndOfFile write FOnEndOfFile;
    property FileName: String read FFileName write SetFileName;
    property FileComment: String read FFileComment write FFileComment;
    // Deprecated. Use Entries.AddFileEntry(FileName) or Entries.AddFileEntries(List) instead.
    property Files: TStrings read FFiles; // deprecated;
    property InMemSize: Int64 read FInMemSize write FInMemSize;
    property Entries: TZipFileEntries read FEntries write SetEntries;
  end;

  { TUnZipper }

  TUnZipper = Class(TObject)
  private
    FOnCloseInputStream: TCustomInputStreamEvent;
    FOnCreateStream: TOnCustomStreamEvent;
    FOnDoneStream: TOnCustomStreamEvent;
    FOnOpenInputStream: TCustomInputStreamEvent;
    FUnZipping  : Boolean;
    FBufSize    : LongWord;
    FFileName   : String;         { Name of resulting Zip file                 }
    FOutputPath : String;
    FFileComment: String;
    FEntries    : TFullZipFileEntries;
    FFiles      : TStrings;
    //FZipStream  : TStream;     { I/O file variables                         }

    FOnPercent  : LongInt;
    FOnProgress : TProgressEvent;
    FOnEndOfFile : TOnEndOfFileEvent;
    FOnStartFile : TOnStartFileEvent;
    {$ifdef KAZIP}procedure ReadKaZip(AExtract: Boolean);{$endif}
    {$ifdef JCL7Z}procedure ReadJCL7Z(AExtract: Boolean);{$endif}
    {$ifdef XE2ZIP}procedure ReadXE2Zip(AExtract: Boolean);{$endif}
  protected
    procedure ReadZip(AExtract: Boolean);
    procedure UnZipOneFile(Item: TFullZipFileEntry); virtual;
    procedure SetBufSize(Value: LongWord);
    procedure SetFileName(Value: String);
    procedure SetOutputPath(Value: String);
  public
    constructor Create;
    destructor Destroy; override;
    procedure UnZipAllFiles; overload;
    procedure UnZipFiles(AFileName: String; FileList: TStrings); overload;
    procedure UnZipFiles(FileList: TStrings); overload;
    procedure UnZipAllFiles(AFileName: String); overload;
    procedure Clear;
    procedure Examine;
  public
    property BufferSize: LongWord Read FBufSize write SetBufSize;
    property OnOpenInputStream: TCustomInputStreamEvent read FOnOpenInputStream write FOnOpenInputStream;
    property OnCloseInputStream: TCustomInputStreamEvent read FOnCloseInputStream write FOnCloseInputStream;
    { before unzipping file, you must provide TStream for file content }
    property OnCreateStream: TOnCustomStreamEvent read FOnCreateStream write FOnCreateStream;
    { after unzipping file, you can read file content from TStream
      NOTE! Some decoders do not trigger OnCreateStream and create their own TStream }
    property OnDoneStream: TOnCustomStreamEvent read FOnDoneStream write FOnDoneStream;
    property OnPercent: Integer read FOnPercent write FOnPercent;
    property OnProgress: TProgressEvent read FOnProgress write FOnProgress;
    property OnStartFile: TOnStartFileEvent read FOnStartFile write FOnStartFile;
    property OnEndFile: TOnEndOfFileEvent read FOnEndOfFile write FOnEndOfFile;
    property FileName: String read FFileName write SetFileName;
    property OutputPath: String read FOutputPath write SetOutputPath;
    property FileComment: String read FFileComment;
    property Files: TStrings read FFiles;
    property Entries: TFullZipFileEntries read FEntries;
  end;


implementation

uses
  zearchhelper;

const
  DirectorySeparator = PathDelim;
  DefaultInMemSize   = 256*1024;   { Files larger than 256k are processed on disk   }
  DefaultBufSize     = 16384;      { Use 16K file buffers                             }

  SErrBufsizeChange = 'Changing buffer size is not allowed while (un)zipping.';
  SErrFileChange = 'Changing output file name is not allowed while (un)zipping.';
  SErrInvalidCRC = 'Invalid CRC checksum while unzipping %s.';
  SErrCorruptZIP = 'Corrupt ZIP file %s.';
  SErrUnsupportedCompressionFormat = 'Unsupported compression format %d';
  SErrUnsupportedMultipleDisksCD = 'A central directory split over multiple disks is unsupported.';
  SErrMaxEntries = 'Encountered %d file entries; maximum supported is %d.';
  SErrMissingFileName = 'Missing filename in entry %d.';
  SErrMissingArchiveName = 'Missing archive filename in streamed entry %d.';
  SErrFileDoesNotExist = 'File "%s" does not exist.';
  SErrFileTooLarge = 'File size %d is larger than maximum supported size %d.';
  SErrPosTooLarge = 'Position/offset %d is larger than maximum supported %d.';
  SErrNoFileName = 'No archive filename for examine operation.';
  SErrNoStream = 'No stream is opened.';
  SErrEncryptionNotSupported = 'Cannot unzip item "%s" : encryption is not supported.';
  SErrPatchSetNotSupported = 'Cannot unzip item "%s" : Patch sets are not supported.';
  SErrNoCompressor = 'Compressor not defined.';

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

{ TZipFileEntry }

procedure TZipFileEntry.Assign(Source: TPersistent);
var
  Z: TZipFileEntry;
begin
  if Source is TZipFileEntry then
  begin
    Z := Source as TZipFileEntry;
    FArchiveFileName := Z.FArchiveFileName;
    FDiskFileName := Z.FDiskFileName;
    FSize := Z.FSize;
    FDateTime := Z.FDateTime;
    FStream := Z.FStream;
    FOS := Z.OS;
    FAttributes := Z.Attributes;
  end
  else
    inherited Assign(Source);
end;

constructor TZipFileEntry.Create(ACollection: TCollection);
begin
  inherited Create(ACollection);
  FCompressionLevel := clDefault;
  FDateTime := Now;
  FAttributes := 0;
end;

function TZipFileEntry.GetArchiveFileName: String;
begin
  Result := FArchiveFileName;
  If (Result = '') then
    Result := FDiskFileName;
end;

function TZipFileEntry.IsDirectory: Boolean;
begin
  Result := (DiskFileName <> '') and (DiskFileName[Length(DiskFileName)] = DirectorySeparator);
end;

function TZipFileEntry.IsLink: Boolean;
begin
  Result := False;
end;

procedure TZipFileEntry.SetArchiveFileName(const AValue: String);
begin
  if FArchiveFileName = AValue then Exit;
  // Zip standard: filenames inside the zip archive have / path separator
  if DirectorySeparator = '/' then
    FArchiveFileName := AValue
  else
    FArchiveFileName := StringReplace(AValue, DirectorySeparator, '/', [rfReplaceAll]);
end;

procedure TZipFileEntry.SetDiskFileName(const AValue: String);
begin
  if FDiskFileName = AValue then Exit;
  // Zip file uses / as directory separator on all platforms
  // so convert to separator used on current OS
  if DirectorySeparator = '/' then
    FDiskFileName := AValue
  else
    FDiskFileName := StringReplace(AValue, '/', DirectorySeparator, [rfReplaceAll]);
end;

{ TZipFileEntries }

procedure TZipFileEntries.AddFileEntries(const List: TStrings);
var
  i: Integer;
begin
  for i:=0 to List.Count-1 do
    AddFileEntry(List[i]);
end;

function TZipFileEntries.AddFileEntry(const AStream: TSTream;
  const AArchiveFileName: String): TZipFileEntry;
begin
  Result := Add as TZipFileEntry;
  Result.Stream := AStream;
  Result.ArchiveFileName := AArchiveFileName;
end;

function TZipFileEntries.AddFileEntry(const ADiskFileName,
  AArchiveFileName: String): TZipFileEntry;
begin
  Result := AddFileEntry(ADiskFileName);
  Result.ArchiveFileName := AArchiveFileName;
end;

function TZipFileEntries.AddFileEntry(const ADiskFileName: String): TZipFileEntry;
begin
  Result := Add as TZipFileEntry;
  Result.DiskFileName := ADiskFileName;
end;

function TZipFileEntries.GetZ(AIndex: Integer): TZipFileEntry;
begin
  Result := TZipFileEntry(Items[AIndex]);
end;

procedure TZipFileEntries.SetZ(AIndex: Integer; const AValue: TZipFileEntry);
begin
  Items[AIndex] := AValue;
end;

{ TFullZipFileEntries }

function TFullZipFileEntries.GetFZ(AIndex: Integer): TFullZipFileEntry;
begin
  Result := TFullZipFileEntry(Items[AIndex]);
end;

procedure TFullZipFileEntries.SetFZ(AIndex: Integer;
  const AValue: TFullZipFileEntry);
begin
  Items[AIndex] := AValue;
end;

{ TZipper }

constructor TZipper.Create;
begin
  inherited;
  FBufSize := DefaultBufSize;
  FFiles := TStringList.Create();
  TStringlist(FFiles).Sorted := True;
  FEntries := TFullZipFileEntries.Create(TFullZipFileEntry);
  FOnPercent := 1;
end;

destructor TZipper.Destroy;
begin
  Clear();
  FreeAndNil(FEntries);
  FreeAndNil(FFiles);
  inherited;
end;

procedure TZipper.BuildZipDirectory;
begin
{$if Defined(XE2ZIP)}
  BuildZipDirectoryXE2Zip();
{$elseif Defined(KAZIP)}
  BuildZipDirectoryKaZip();
{$elseif Defined(JCL7Z)}
  BuildZipDirectoryJCL7Z();
{$elseif Defined(ABZIP)}
  BuildZipDirectoryAbZip();
{$elseif Defined(SCIZIP)}
  BuildZipDirectorySciZip();
{$else}
  raise EZipError.Create(SErrNoCompressor);
{$ifend}
end;

function TZipper.CheckEntries: Integer;
var
  i: Integer;
begin
  for i:=0 to FFiles.Count-1 do
    FEntries.AddFileEntry(FFiles[i]);
  Result := FEntries.Count;
end;

procedure TZipper.Clear;
begin
  FEntries.Clear();
  FFiles.Clear();
end;

procedure TZipper.GetFileInfo();
var
  Item: TZipFileEntry;
  Info: TSearchRec;
  i: Integer;
begin
  for i := 0 to FEntries.Count-1 do
  begin
    Item := FEntries[i];
    if Item.Stream = nil then
    begin
      if (Item.DiskFileName = '') then
        raise EZipError.CreateFmt(SErrMissingFileName, [i]);
      if FindFirst(Item.DiskFileName, (faAnyFile + faDirectory), Info) = 0 then
      begin
        try
          Item.Size := Info.Size;
          Item.DateTime := FileDateToDateTime(Info.Time);
          Item.Attributes := Info.Attr;
        finally
          FindClose(Info);
        end;
      end
      else
        raise EZipError.CreateFmt(SErrFileDoesNotExist, [Item.DiskFileName]);
    end
    else
    begin
      if (Item.ArchiveFileName = '') then
        raise EZipError.CreateFmt(SErrMissingArchiveName, [i]);
      Item.Size := Item.Stream.Size;
      if (Item.Attributes = 0) then
      begin
        Item.Attributes := faArchive;
      end;
    end;
  end;
end;

procedure TZipper.SaveToFile(AFileName: string);
var
  i: Integer; //could be qword but limited by FEntries.Count
begin
  FFileName := AFileName;
  if CheckEntries = 0 then
    Exit;
  FZipping := True;
  try
    GetFileInfo(); //get info on file entries in zip
    if FEntries.Count > 0 then
      BuildZipDirectory();
  finally
    FZipping := False;
    // Remove entries that have been added by CheckEntries from Files.
    for i:=0 to FFiles.Count-1 do
      FEntries.Delete(FEntries.Count-1);
  end;
end;

procedure TZipper.SaveToStream(AStream: TStream);
begin
  raise EZipError.Create('Not supported');
end;

procedure TZipper.SetBufSize(Value: LongWord);
begin
  if FZipping then
    raise EZipError.Create(SErrBufsizeChange);
  if Value >= DefaultBufSize then
    FBufSize := Value;
end;

procedure TZipper.SetEntries(const AValue: TZipFileEntries);
begin
  if FEntries = AValue then Exit;
  FEntries.Assign(AValue);
end;

procedure TZipper.SetFileName(Value: String);
begin
  if FZipping then
    raise EZipError.Create(SErrFileChange);
  FFileName := Value;
end;

procedure TZipper.ZipAllFiles;
begin
  SaveToFile(FileName);
end;

procedure TZipper.ZipFiles(Entries: TZipFileEntries);
begin
  FEntries.Assign(Entries);
  ZipAllFiles();
end;

procedure TZipper.ZipFiles(AFileName: String; Entries: TZipFileEntries);
begin
  FFileName := AFileName;
  ZipFiles(Entries);
end;

procedure TZipper.ZipFiles(AFileName: String; FileList: TStrings);
begin
  FFileName := AFileName;
  ZipFiles(FileList);
end;

procedure TZipper.ZipFiles(FileList: TStrings);
begin
  FFiles.Assign(FileList);
  ZipAllFiles();
end;

{$ifdef KAZIP}
procedure TZipper.BuildZipDirectoryKaZip();
var
  zip: TKaZip;
  i: Integer;
  Item: TZipFileEntry;
begin
  zip := TKaZip.Create(nil);
  try
    zip.CreateZip(FileName);
    zip.Open(FileName);
    zip.StoreFolders := False;
    for i := 0 to Entries.Count-1 do
    begin
      Item := Entries[i];
      if Assigned(Item.Stream) then
      begin
        Item.Stream.Position := 0;
        zip.AddStream(Item.ArchiveFileName, Item.Stream);
      end
      else
      begin
        zip.AddFile(Item.DiskFileName, Item.ArchiveFileName);
      end;
    end;
    zip.Close();
  finally
    zip.Free();
  end;
end;
{$endif}

{$ifdef JCL7Z}
procedure TZipper.BuildZipDirectoryJCL7Z();
var
  zip: TJclZipCompressArchive;
  i: Integer;
  Item: TZipFileEntry;
begin
  zip := TJclZipCompressArchive.Create(FileName);
  try
    for i := 0 to Entries.Count-1 do
    begin
      Item := Entries[i];
      if Assigned(Item.Stream) then
      begin
        Item.Stream.Position := 0;
        zip.AddFile(Item.ArchiveFileName, Item.Stream, False);
      end
      else
      begin
        zip.AddFile(Item.ArchiveFileName, Item.DiskFileName);
      end;
    end;
    zip.Compress();
  finally
    FreeAndNil(zip);
  end;
end;
{$endif}

{$ifdef ABZIP}
procedure TZipper.BuildZipDirectoryAbZip();
var
  zip: TAbZipper;
  i: Integer;
  Item: TZipFileEntry;
  TmpStream: TFileStream;
begin
  zip := TAbZipper.Create(nil);
  try
    zip.ArchiveType := atZip;
    zip.ForceType := True;
    zip.FileName := FileName;
    //zip.BaseDirectory := FTmpZipFolder;
    zip.StoreOptions := [soRecurse];
    //zip.AddFiles('*', faAnyFile);
    for i := 0 to Entries.Count-1 do
    begin
      Item := Entries[i];
      if Assigned(Item.Stream) then
      begin
        Item.Stream.Position := 0;
        zip.AddFromStream(Item.ArchiveFileName, Item.Stream);
      end
      else
      begin
        TmpStream := TFileStream.Create(Item.DiskFileName, fmOpenRead);
        try
          zip.AddFromStream(Item.ArchiveFileName, TmpStream);
        finally
          TmpStream.Free();
        end;
      end;
    end;
    zip.Save();
    zip.CloseArchive();
  finally
    zip.Free();
  end;
end;
{$endif}


{$ifdef SCIZIP}
procedure TZipper.BuildZipDirectorySciZip();
var
  zip: TZipFile;
  i: integer;
  s: string;
  sl: TStringList;
  stream: TFileStream;
  buffer: AnsiString;
  Item: TZipFileEntry;
begin
  s := ExtractFilePath(FileName);
  if not ForceDirectories(s) then
  begin
    Exit;
  end;

  zip := TZipFile.Create;
  try
    for i := 0 to Entries.Count-1 do
    begin
      Item := Entries[i];
      if Assigned(Item.Stream) then
      begin
        if Item.Stream.Size = 0 then
          Continue;
        zip.AddFile(AnsiString(Item.ArchiveFileName));
        Item.Stream.Position := 0;
        SetLength(buffer, Item.Stream.Size);
        Item.Stream.ReadBuffer(buffer[1], Item.Stream.Size);
        zip.Data[zip.Count - 1] := buffer;
      end
      else
      begin
        TmpStream := TFileStream.Create(Item.DiskFileName, fmOpenRead);
        try
          zip.AddFile(AnsiString(Item.ArchiveFileName));
          TmpStream.Position := 0;
          SetLength(buffer, TmpStream.Size);
          TmpStream.ReadBuffer(buffer[1], TmpStream.Size);
          zip.Data[zip.Count - 1] := buffer;
        finally
          TmpStream.Free();
        end;
      end;
    end;

    zip.SaveToFile(FileName);
  finally
    zip.Free;
  end;
end;
{$endif}


{$ifdef XE2ZIP}
procedure TZipper.BuildZipDirectoryXE2Zip();
var
  zip: TZipFile;
  s: string;
  i: Integer;
  Item: TZipFileEntry;
begin
  s := ExtractFilePath(FileName);
  if (not ForceDirectories(s)) then
    Exit;

  zip := TZipFile.Create();
  try
    zip.Open(FileName, zmWrite);
    for i := 0 to Entries.Count-1 do
    begin
      Item := Entries[i];
      if Assigned(Item.Stream) then
      begin
        Item.Stream.Position := 0;
        zip.Add(Item.Stream, Item.ArchiveFileName);
      end
      else
      begin
        zip.Add(Item.DiskFileName, Item.ArchiveFileName);
      end;
    end;
    zip.Close();
  finally
    FreeAndNil(zip);
  end;
end;
{$endif}

{ TUnZipper }

constructor TUnZipper.Create;
begin
  inherited;
  FBufSize := DefaultBufSize;
  FFiles := TStringList.Create();
  TStringlist(FFiles).Sorted := True;
  FEntries := TFullZipFileEntries.Create(TFullZipFileEntry);
  FOnPercent := 1;
end;

destructor TUnZipper.Destroy;
begin
  Clear();
  FreeAndNil(FFiles);
  FreeAndNil(FEntries);
  inherited;
end;

procedure TUnZipper.Clear;
begin
  FFiles.Clear();
  FEntries.Clear();
end;

procedure TUnZipper.Examine;
begin
  if (not Assigned(FOnOpenInputStream)) and (FFileName = '') then
    Raise EZipError.Create(SErrNoFileName);

  ReadZip(False);
end;

procedure TUnZipper.SetBufSize(Value: LongWord);
begin
  if FUnZipping then
    raise EZipError.Create(SErrBufsizeChange);
  if Value >= DefaultBufSize then
    FBufSize := Value;
end;

procedure TUnZipper.SetFileName(Value: String);
begin
  if FUnZipping then
    raise EZipError.Create(SErrFileChange);
  FFileName := Value;
end;

procedure TUnZipper.SetOutputPath(Value: String);
begin
  if FUnZipping then
    raise EZipError.Create(SErrFileChange);
  FOutputPath := Value;
end;

procedure TUnZipper.UnZipAllFiles;
begin
  ReadZip(True);
end;

procedure TUnZipper.UnZipAllFiles(AFileName: String);
begin
  FFileName := AFileName;
  UnZipAllFiles();
end;

procedure TUnZipper.UnZipFiles(FileList: TStrings);
begin
  FFiles.Assign(FileList);
  UnZipAllFiles();
end;

procedure TUnZipper.UnZipFiles(AFileName: String; FileList: TStrings);
begin
  FFileName := AFileName;
  UNzipFiles(FileList);
end;

procedure TUnZipper.UnZipOneFile(Item: TFullZipFileEntry);
begin

end;

procedure TUnZipper.ReadZip(AExtract: Boolean);
begin
  FUnZipping := True;
  try
    Entries.Clear();
    {$if Defined(XE2ZIP)}
    ReadXE2Zip(AExtract);
    {$elseif Defined(KAZIP)}
    ReadKaZip(AExtract);
    {$elseif Defined(JCL7Z)}
    ReadJCL7Z(AExtract);
    {$elseif Defined(ABZIP)}
    ReadAbZip(AExtract);
    {$elseif Defined(SCIZIP)}
    ReadSciZip(AExtract);
    {$else}
    raise EZipError.Create(SErrNoCompressor);
    {$ifend}
  finally
    FUnZipping := False;
  end;
end;

{$ifdef KAZIP}
procedure TUnZipper.ReadKaZip(AExtract: Boolean);
var
  zip: TKaZip;
  i: Integer;
  Item: TZipFileEntry;
  ZipItem: TKAZipEntriesEntry;
  TmpStream: TStream;
begin
  zip := TKaZip.Create(nil);
  try
    zip.Open(FileName);
    for i := 0 to zip.Entries.Count-1 do
    begin
      ZipItem := zip.Entries.Items[i];

      Item := Entries.AddFileEntry('', ZipItem.FileName);
      if AExtract and ((FFiles.Count = 0) or (FFiles.IndexOf(Item.ArchiveFileName) <> -1)) then
      begin
        if Assigned(OnCreateStream) and Assigned(OnDoneStream) then
        begin
          OnCreateStream(Self, TmpStream, (Item as TFullZipFileEntry));
          try
            zip.ExtractToStream(ZipItem, TmpStream);
          finally
            OnDoneStream(Self, TmpStream, (Item as TFullZipFileEntry));
          end;
        end;
      end;
    end;
    zip.Close();
  finally
    zip.Free();
  end;
end;
{$endif}

{$ifdef JCL7Z}
procedure TUnZipper.ReadJCL7Z(AExtract: Boolean);
var
  zip: TJclZipDecompressArchive;
  i: Integer;
  Item: TZipFileEntry;
  ZipItem: TJclCompressionItem;
  TmpStream: TStream;
begin
  zip := TJclZipDecompressArchive.Create(FileName);
  try
    zip.ListFiles();
    for i := 0 to zip.ItemCount-1 do
    begin
      ZipItem := zip.Items[i];

      Item := Entries.AddFileEntry('', ZipItem.PackedName);
      if AExtract and ((FFiles.Count = 0) or (FFiles.IndexOf(Item.ArchiveFileName) <> -1)) then
      begin
        if Assigned(OnCreateStream) and Assigned(OnDoneStream) then
        begin
          OnCreateStream(Self, TmpStream, (Item as TFullZipFileEntry));
          try
            ZipItem.Stream := TmpStream;
            ZipItem.Selected := True;
            zip.ExtractSelected();
            ZipItem.Stream := nil;
            ZipItem.Selected := False;
          finally
            OnDoneStream(Self, TmpStream, (Item as TFullZipFileEntry));
          end;
        end;
      end;
    end;
  finally
    zip.Free();
  end;
end;
{$endif}

{$ifdef XE2ZIP}
procedure TUnZipper.ReadXE2Zip(AExtract: Boolean);
var
  zip: TZipFile;
  i: Integer;
  Item: TZipFileEntry;
  TmpStream: TStream;
  localHeader: TZipHeader;
begin
  zip := TZipFile.Create();
  try
    zip.Open(FileName, zmRead);
    for i := 0 to zip.FileCount-1 do
    begin
      Item := Entries.AddFileEntry('', zip.FileNames[i]);
      if AExtract and ((FFiles.Count = 0) or (FFiles.IndexOf(Item.ArchiveFileName) <> -1)) then
      begin
        if Assigned(OnCreateStream) and Assigned(OnDoneStream) then
        begin
          // stream created inside zip.Read() and destroyed outside
          //OnCreateStream(Self, TmpStream, (Item as TFullZipFileEntry));
          try
            zip.Read(i, TmpStream, localHeader);
          finally
            OnDoneStream(Self, TmpStream, (Item as TFullZipFileEntry));
          end;
        end;
      end;
    end;
    zip.Close();
  finally
    FreeAndNil(zip);
  end;
end;
{$endif}


end.
