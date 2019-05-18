(*
Delphi freeware Zip Project
Copyright (c) 2005 by Kiril Antonov

=== DESCRIPTTION ===
KAZIP is fast, simple zip archiver/dearchiver which uses most popular
Inflate/Deflate zip compression format. KAZip is totaly based on Delphi
VCL - no dll, ActiveX or other external libraries. KAZip is totaly stream
oriented so you can deal with data only in memory without creating temporary
files, etc. If you need to add zip/unzip functionality to your application,
KAZIP is the right solution.

=== INSTALLATION ===
If you want to enable BZIP2 functionality:
- Download BZIP2 pas files and put them in the KAZIP folder or folder in the delphi search path
- Open KAZip.Pas and change {DEFINE USE_BZIP2} to {$DEFINE USE_BZIP2} Open KZ.DPK and press Install

=== LICENCE ===
KAZIP IS FREE FOR COMMERICAL AND NON COMMERICAL USE! if you want to support
free coding purchase KADBZIP - a database-oriented version od KAZip

=== DISCLAIMER OF WARRANTY ===
COMPONENTS ARE SUPPLIED "AS IS" WITHOUT WARRANTY OF ANY KIND. THE AUTHOR
DISCLAIMS ALL WARRANTIES, EXPRESSED OR IMPLIED, INCLUDING, WITHOUT LIMITATION,
THE WARRANTIES OF MERCHANTABILITY AND OF FITNESS FOR ANY PURPOSE. THE AUTHOR
ASSUMES NO LIABILITY FOR DAMAGES, DIRECT OR CONSEQUENTIAL, WHICH MAY RESULT
FROM THE USE OF COMPONENTS. USE THIS COMPONENTS AT YOUR OWN RISK

For contacts:
my e-mail: kiril at pari.bg, kirila at abv.bg
my site: http://kadao.dir.bg/

Best regards Kiril Antonov Sofia Bulgaria

*)
unit KAZip;

{
 How to create a zip file:

  var
    zip: TKaZip;
    zipFile: string;
  begin
    zipFile := 'C:\DelphiComponents\KaZip\Component\TestArchive.zip';

    zip := TKaZip.Create(nil);
    zip.CreateZip(zipFile);
    zip.Open(zipFile);
    zip.AddFile('C:\DelphiComponents\KaZip\Component\KAZip.pas', 'kazip.pas');
    zip.Close;
    zip.Free;

 How to add a stream to a zip file:
  var
    zip: TKaZip;
    zipFile: string;
  begin
    zipFile := 'C:\DelphiComponents\KaZip\Component\TestArchive.zip';

    zip := TKaZip.Create(nil);
    zip.CreateZip(zipFile);
    zip.Open(zipFile);
    zip.AddStream('Attachment1.jpg', stream);
    zip.Close;
   zip.Free;

 How to add a string to a zip file:
  var
    zip: TKaZip;
    zipFile: string;
    data: string;
    ms: TStringStream;
  begin
    zipFile := 'C:\DelphiComponents\KaZip\Component\TestArchive.zip';

    zip := TKaZip.Create(nil);
    zip.CreateZip(zipFile);
    zip.Open(zipFile);

    ms := TStringStream.Create('The quick brown fox jumped over the lazy dog then sat on a log');
   try
     zip.AddStream('New Document.txt', stream);
   finally
     ms.Free;
   end;
   zip.Close;
   zip.Free;

 How to add a WideString to a zip file (Remember, a widestring is UTF-16 encoded)

  var
    zip: TKaZip;
    zipFile: string;
    ms: TStringStream;
    docxml: WideString;
  const
    UTF16BOM = #$FF#$FE; //little-endian (i.e. Intel) order
  begin
    zipFile := 'C:\DelphiComponents\KaZip\Component\TestArchive.zip';

    zip := TKaZip.Create(nil);
    zip.CreateZip(zipFile);
    zip.Open(zipFile);

    docxml := '<Sample>The quick brown fox jumped over the lazy dog then sat on a log</Sample>';
    ms := TMemoryStream.Create
    try
      ms.Write(UTF16BOM[1], 2); //Write the UTF-16 Byte Order Mark (optional)
      ms.Write(docxml[1], Length(docxml)*2);
      zip.AddStream('Mouse.xml', stream);
    finally
      ms.Free;
    end;
    zip.Close;
    zip.Free;

}

interface

{DEFINE USE_BZIP2}

uses
  Windows, SysUtils, Classes, Masks, TypInfo
  {$IFDEF USE_BZIP2}, BZip2{$ENDIF}
  , ZLibEx //zlib 1.2.8
  ;

type
  TZipSaveMethod = (FastSave, RebuildAll);
  TZipCompressionType = (ctNormal, ctMaximum, ctFast, ctSuperFast, ctNone, ctUnknown);
  TOverwriteAction = (oaSkip, oaSkipAll, oaOverwrite, oaOverwriteAll);

  //Zip Compression Methods. DEFLATE (CM=8) is default
//  TZipCompressionMethod = (cmStored, cmShrunk, cmReduced1, cmReduced2, cmReduced3, cmReduced4, cmImploded, cmTokenizingReserved, cmDeflated, cmDeflated64, cmDCLImploding, cmPKWAREReserved);
  TZipCompressionMethod = Word;
const
  ZipCompressionMethod_Store =       0;  //The file is stored (no compression)
  ZipCompressionMethod_Shrink =     1;  //The file is Shrunk
  ZipCompressionMethod_Reduce1 =     2;  //The file is Reduced with compression factor 1
  ZipCompressionMethod_Reduce2 =     3;  //The file is Reduced with compression factor 2
  ZipCompressionMethod_Reduce3 =     4;  //The file is Reduced with compression factor 3
  ZipCompressionMethod_Reduce4 =     5;  //The file is Reduced with compression factor 4
  ZipCompressionMethod_Implode =     6;  //The file is Imploded
//  ZipCompressionMethod_Tokenizing =   7;  //Reserved for Tokenizing compression algorithm (this method is not used by by PKZIP)
  ZipCompressionMethod_Deflate =     8;  //The file is Deflated
  ZipCompressionMethod_Deflate64 =   9;  //Enhanced Deflating using Deflate64(tm)
//  ZipCompressionMethod_ =         10;  //PKWARE Data Compression Library Imploding
//  ZipCompressionMethod_ =         11;  //Reserved by PKWARE
  ZipCompressionMethod_BZIP2 =       12;  //File is compressed using BZIP2 algorithm
//  ZipCompressionMethod_ =         13; //Reserved by PKWARE
  ZipCompressionMethod_LZMA =       14; //LZMA (EFS)
//  ZipCompressionMethod_ =         15; //Reserved by PKWARE
//  ZipCompressionMethod_ =         16; //Reserved by PKWARE
//  ZipCompressionMethod_ =         17; //Reserved by PKWARE
//  ZipCompressionMethod_IBMTerse =     18; //File is compressed using IBM TERSE (new)
//  ZipCompressionMethod_IBMLZ77 =     19; //IBM LZ77 z Architecture (PFS)
  ZipCompressionMethod_WavPack=     97; //WavPack compressed data
  ZipCompressionMethod_PPMd =       98; //PPMd version I, Rev 1

  { DoChange Events }
  ZIP_CLOSED       = 0; // Zip is Closed;
  ZIP_OPENED       = 1; // Zip is Opened;
  ZIP_ITEM_ADDED   = 2; // Item is added to the zip
  ZIP_ITEM_REMOVED = 3; // Item is removed from the Zip
  ZIP_ITEM_COMMENT = 4; // Item comment changed
  ZIP_ITEM_NAME    = 5; // Item name changed (single item)
  ZIP_ITEMS_NAME   = 6; // Item name changed  (multiple items)

  // for compatibility
  {$ifdef WIN32}
  faReadOnly  = $00000001;
  faHidden    = $00000002;
  faSysFile   = $00000004;
  faArchive   = $00000020;
  {$endif}

type
  TOnDecompressFile = procedure(Sender: TObject; Current, Total: Integer) of object;
  TOnCompressFile = procedure(Sender: TObject; Current, Total: Integer) of object;
  TOnZipOpen = procedure(Sender: TObject; Current, Total: Integer) of object;
  TOnZipChange = procedure(Sender: TObject; ChangeType: Integer) of object;
  TOnAddItem = procedure(Sender: TObject; ItemName: string) of object;
  TOnRebuildZip = procedure(Sender: TObject; Current, Total: Integer) of object;
  TOnRemoveItems = procedure(Sender: TObject; Current, Total: Integer) of object;
  TOnOverwriteFile = procedure(Sender: TObject; var FileName: string; var Action: TOverwriteAction) of object;

  TKAZipEntries = class;
  TKAZip = class;
  TBytes = array of Byte;
  TZipSignature = array[0..3] of AnsiChar;

  TLocalFileHeader = packed record
    LocalFileHeaderSignature: TZipSignature; //    4 bytes  = 0x04034b50 = 'PK'#03#04
    VersionNeededToExtract: WORD;       //    2 bytes
    GeneralPurposeBitFlag: WORD;        //    2 bytes
    CompressionMethod: TZipCompressionMethod; //    2 bytes
    LastModFileTimeDate: LongWord;      //    4 bytes
    Crc32: LongWord;                    //    4 bytes
    CompressedSize: LongWord;           //    4 bytes
    UncompressedSize: LongWord;         //    4 bytes
    FilenameLength: WORD;               //    2 bytes
    ExtraFieldLength: WORD;             //    2 bytes
  end;

  TZipExtraDataRecord = packed record
    Signature: TZipSignature;           //    4 bytes 0x08064b50
    FieldLength: LongWord;              //    4 bytes
  end;

  TDataDescriptor = packed record
    DescriptorSignature: LongWord; //    4 bytes OPTIONAL 0x08074b50
    Crc32: LongWord; //    4 bytes
    CompressedSize: LongWord; //    4 bytes
    UncompressedSize: LongWord; //    4 bytes
  end;

  TDataDescriptor64 = packed record
    DescriptorSignature: LongWord; //    8 bytes OPTIONAL 0x08074b50
    Crc32: LongWord; //    4 bytes
    CompressedSize: Int64; //    8 bytes
    UncompressedSize: Int64; //    8 bytes
  end;

  TCentralDirectoryFile = packed record
    CentralFileHeaderSignature: TZipSignature; //    4 bytes = 0x02014b50 = 'PK'#01#02
    VersionMadeBy: WORD; //    2 bytes
    VersionNeededToExtract: WORD; //    2 bytes
    GeneralPurposeBitFlag: WORD; //    2 bytes
    CompressionMethod: TZipCompressionMethod; //    2 bytes
    LastModFileTimeDate: LongWord; //    4 bytes
    Crc32: LongWord; //    4 bytes
    CompressedSize: LongWord; //    4 bytes
    UncompressedSize: LongWord; //    4 bytes
    FilenameLength: WORD; //    2 bytes
    ExtraFieldLength: WORD; //    2 bytes
    FileCommentLength: WORD; //    2 bytes
    DiskNumberStart: WORD; //    2 bytes
    InternalFileAttributes: WORD; //    2 bytes
    ExternalFileAttributes: LongWord; //    4 bytes
    RelativeOffsetOfLocalHeader: LongWord; //    4 bytes
    //FileName: AnsiString; //    variable size
    //ExtraField: AnsiString; //    variable size
    //FileComment: AnsiString; //    variable size
  end;

  TEndOfCentralDir = packed record
    EndOfCentralDirSignature: TZipSignature; //    4 bytes = 0x06054b50 = 'PK'#05#06
    NumberOfThisDisk: WORD; //    2 bytes
    NumberOfTheDiskWithTheStart: WORD; //    2 bytes
    TotalNumberOfEntriesOnThisDisk: WORD; //    2 bytes
    TotalNumberOfEntries: WORD; //    2 bytes
    SizeOfTheCentralDirectory: LongWord; //    4 bytes
    OffsetOfStartOfCentralDirectory: LongWord; //    4 bytes
    ZipfileCommentLength: WORD; //    2 bytes
  end;

  TKAZipEntriesEntry = class(TCollectionItem)
  private
    //FCentralDirectoryFile: TCentralDirectoryFile;
    //FLocalFileHeader: TLocalFile;
    FParent: TKAZipEntries;
    FVersionMadeBy: Word;
    FVersionNeededToExtract: Word;
    FFileName: AnsiString;
    FFileComment: AnsiString;
    FExtraField: AnsiString;
    FLocalFileEntrySize: LongWord;
    FCentralFileEntrySize: LongWord;
    FCRC32: LongWord;
    FSizeUncompressed: Int64;
    FSizeCompressed: Int64;
    FDiskNumber: Word;
    FInternalFileAttributes: Word;
    FExternalFileAttributes: LongWord;
    FLocalFileHeaderOffset: Int64;
    FGeneralPurposeBitFlag: Word;
    FCompressionMethod: TZipCompressionMethod;
    FIsEncrypted: Boolean;
    FIsFolder: Boolean;
    FDate: TDateTime;
    FCompressionType: TZipCompressionType;
    FSelected: Boolean;
    FCompressedDataPos: Int64;

    procedure SetSelected(const Value: Boolean);
    function GetLocalEntrySize: LongWord;
    function GetCentralEntrySize: LongWord;
    function GetLocalHeadSize: LongWord;
    procedure SetComment(const Value: AnsiString);
    procedure SetFileName(const Value: AnsiString);
    procedure UpdateProperties();
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;

    function ParseFromCentralDirectory(MS: TStream): Boolean;
    function ParseFromLocalHeader(MS: TStream): Boolean;
    procedure FillCentralDirectory(out CDFile: TCentralDirectoryFile);
    procedure FillLocalHeader(out LFH: TLocalFileHeader);

    function GetCompressedData: AnsiString; overload;
    function GetCompressedData(Stream: TStream): Integer; overload;
    procedure ExtractToFile(FileName: string);
    procedure ExtractToStream(Stream: TStream);
    procedure SaveToFile(FileName: string);
    procedure SaveToStream(Stream: TStream);
    function Test: Boolean;

    property Version: Word read FVersionMadeBy;
    property VersionMin: Word read FVersionNeededToExtract;
    property BitFlag: Word read FGeneralPurposeBitFlag;
    property Date: TDateTime read FDate write FDate;
    property CRC32: LongWord read FCRC32 write FCRC32;
    property SizeUncompressed: Int64 read FSizeUncompressed write FSizeUncompressed;
    property SizeCompressed: Int64 read FSizeCompressed write FSizeCompressed;
    property FileName: AnsiString read FFileName write SetFileName;
    property ExtraField: AnsiString read FExtraField write FExtraField;
    property Comment: AnsiString read FFileComment write SetComment;
    property DiskNumber: Word read FDiskNumber;
    property InternalAttributes: Word read FInternalFileAttributes;
    property Attributes: LongWord read FExternalFileAttributes write FExternalFileAttributes;
    property LocalOffset: Int64 read FLocalFileHeaderOffset write FLocalFileHeaderOffset;
    property IsEncrypted: Boolean read FIsEncrypted;
    property IsFolder: Boolean read FIsFolder;
    property CompressionMethod: TZipCompressionMethod read FCompressionMethod write FCompressionMethod;
    property CompressionType: TZipCompressionType read FCompressionType;
    property LocalEntrySize: LongWord read GetLocalEntrySize;
    property CentralEntrySize: LongWord read GetCentralEntrySize;
    property Selected: Boolean read FSelected write SetSelected;
    property CompressedDataPos: Int64 read FCompressedDataPos;
  end;

  TKAZipEntries = class(TCollection)
  private
    { Private declarations }
    FParent: TKAZip;

    function GetHeaderEntry(Index: Integer): TKAZipEntriesEntry;
    procedure SetHeaderEntry(Index: Integer; const Value: TKAZipEntriesEntry);
  protected
    { Protected declarations }
    procedure SortList(List: TList);
    function FileTime2DateTime(FileTime: TFileTime): TDateTime;
    //**************************************************************************
  public
    { Public declarations }
    procedure ParseZip(MS: TStream);
    constructor Create(AOwner: TKAZip; MS: TStream); overload;
    constructor Create(AOwner: TKAZip); overload;
    destructor Destroy; override;
    //**************************************************************************
    function IndexOf(const FileName: string): Integer;
    //**************************************************************************
    function AddFile(FileName, NewFileName: string): TKAZipEntriesEntry; overload;
    function AddFile(FileName: string): TKAZipEntriesEntry; overload;
    function AddFiles(FileNames: TStrings): Boolean;
    function AddFolder(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
    function AddFilesAndFolders(FileNames: TStrings; RootFolder: string; WithSubFolders: Boolean): Boolean;
    function AddStream(FileName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry; overload;
    function AddStream(FileName: string; Stream: TStream): TKAZipEntriesEntry; overload;
    function AddEntryThroughStream(FileName: string; FileDate: TDateTime; FileAttr: Word): TStream;

    //**************************************************************************
    procedure Remove(ItemIndex: Integer); overload;
    procedure Remove(Item: TKAZipEntriesEntry); overload;
    procedure Remove(FileName: string); overload;
    procedure RemoveFiles(List: TList);
    procedure RemoveSelected;
    procedure Rebuild;
    //**************************************************************************
    procedure Select(WildCard: string);
    procedure SelectAll;
    procedure DeSelectAll;
    procedure InvertSelection;
    //**************************************************************************
    procedure Rename(Item: TKAZipEntriesEntry; NewFileName: string); overload;
    procedure Rename(ItemIndex: Integer; NewFileName: string); overload;
    procedure Rename(FileName: string; NewFileName: string); overload;
    procedure CreateFolder(FolderName: string; FolderDate: TDateTime);
    procedure RenameFolder(FolderName: string; NewFolderName: string);
    procedure RenameMultiple(Names: TStringList; NewNames: TStringList);

    //**************************************************************************
    procedure ExtractToFile(Item: TKAZipEntriesEntry; FileName: string); overload;
    procedure ExtractToFile(ItemIndex: Integer; FileName: string); overload;
    procedure ExtractToFile(FileName, DestinationFileName: string); overload;
    procedure ExtractToStream(Item: TKAZipEntriesEntry; TargetStream: TStream);
    procedure ExtractAll(TargetDirectory: string);
    procedure ExtractSelected(TargetDirectory: string);
    //**************************************************************************
    property Items[Index: Integer]: TKAZipEntriesEntry read GetHeaderEntry write SetHeaderEntry;
  end;

  TKAZip = class(TComponent)
  private
    { Private declarations }
    FZipHeader: TKAZipEntries;
    FIsDirty: Boolean;
    FEndOfCentralDirPos: Cardinal;
    FEndOfCentralDir: TEndOfCentralDir;

    FZipCommentPos: LongWord;
    FZipComment: TStringList;

    FRebuildECDP: Cardinal;
    FRebuildCP: Cardinal;

    FLocalHeaderNumFiles: Integer;
    FIsZipFile: Boolean;
    FHasBadEntries: Boolean;
    FFileName: string;
    FFileNames: TStringList;
    FZipSaveMethod: TZipSaveMethod;

    FExternalStream: Boolean;
    FStoreRelativePath: Boolean;
    FZipCompressionType: TZipCompressionType;

    FCurrentDFS: Cardinal;
    FOnDecompressFile: TOnDecompressFile;
    FOnCompressFile: TOnCompressFile;
    FOnZipChange: TOnZipChange;
    FBatchMode: Boolean;

    FOnZipOpen: TOnZipOpen;
    FUseTempFiles: Boolean;
    FStoreFolders: Boolean;
    FOnAddItem: TOnAddItem;
    FComponentVersion: string;
    FOnRebuildZip: TOnRebuildZip;
    FOnRemoveItems: TOnRemoveItems;
    FOverwriteAction: TOverwriteAction;
    FOnOverwriteFile: TOnOverwriteFile;
    FReadOnly: Boolean;
    FApplyAttributes: Boolean;

    procedure SetFileName(const Value: string);
    procedure SetIsZipFile(const Value: Boolean);
    function GetComment: TStrings;
    procedure SetComment(const Value: TStrings);
    procedure SetZipSaveMethod(const Value: TZipSaveMethod);
    procedure SetActive(const Value: Boolean);
    procedure SetZipCompressionType(const Value: TZipCompressionType);
    function GetFileNames: TStrings;
    procedure SetFileNames(const Value: TStrings);
    procedure SetUseTempFiles(const Value: Boolean);
    procedure SetStoreFolders(const Value: Boolean);
    procedure SetOnAddItem(const Value: TOnAddItem);
    procedure SetComponentVersion(const Value: string);
    procedure SetOnRebuildZip(const Value: TOnRebuildZip);
    procedure SetOnRemoveItems(const Value: TOnRemoveItems);
    procedure SetOverwriteAction(const Value: TOverwriteAction);
    procedure SetOnOverwriteFile(const Value: TOnOverwriteFile);
    procedure SetReadOnly(const Value: Boolean);
    procedure SetApplyAtributes(const Value: Boolean);
  protected
    FZipStream: TStream;
    //**************************************************************************
    procedure ParseZip(MS: TStream);
    procedure LoadFromFile(FileName: string);
    procedure LoadFromStream(MS: TStream);
    //**************************************************************************
    procedure RebuildLocalFiles(MS: TStream);
    procedure RebuildCentralDirectory(MS: TStream);
    //**************************************************************************
    procedure OnDecompress(Sender: TObject);
    procedure OnCompress(Sender: TObject);
    procedure DoChange(Sender: TObject; const ChangeType: Integer); virtual;
    //**************************************************************************
    function FindCentralDirectory(MS: TStream): Boolean;
    function ParseCentralHeaders(MS: TStream): Boolean;
    //function GetLocalEntry(MS: TStream; Offset: Integer; HeaderOnly: Boolean): TLocalFile;
    procedure LoadLocalHeaders(MS: TStream);
    function ParseLocalHeaders(MS: TStream): Boolean;
    //**************************************************************************
    function AddStreamFast(ItemName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry; overload;
    function AddStreamRebuild(ItemName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
    function AddFolderChain(ItemName: string): Boolean; overload;
    function AddFolderChain(ItemName: string; FileAttr: Word; FileDate: TDateTime): Boolean; overload;
    function AddFolderEx(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
    //**************************************************************************
    procedure Remove(ItemIndex: Integer; Flush: Boolean); overload;
    procedure RemoveBatch(Files: TList);
    procedure InternalExtractToFile(Item: TKAZipEntriesEntry; FileName: string);
    //**************************************************************************
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    //**************************************************************************
    function GetDelphiTempFileName: string;
    function GetFileName(S: string): string;
    function GetFilePath(S: string): string;
    //**************************************************************************
    procedure CreateZip(Stream: TStream); overload;
    procedure CreateZip(FileName: string); overload;
    procedure Open(FileName: string); overload;
    procedure Open(MS: TStream); overload;
    procedure SaveToStream(Stream: TStream);
    procedure Rebuild;
    procedure FixZip(MS: TStream);
    procedure Close;
    //**************************************************************************
    function AddFile(FileName, NewFileName: string): TKAZipEntriesEntry; overload;
    function AddFile(FileName: string): TKAZipEntriesEntry; overload;
    function AddFiles(FileNames: TStrings): Boolean;
    function AddFolder(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
    function AddFilesAndFolders(FileNames: TStrings; RootFolder: string; WithSubFolders: Boolean): Boolean;
    function AddStream(FileName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry; overload;
    function AddStream(FileName: string; Stream: TStream): TKAZipEntriesEntry; overload;

    function AddEntryThroughStream(FileName: string): TStream; overload;
    function AddEntryThroughStream(FileName: string; FileDate: TDateTime; FileAttr: Word): TStream; overload;

    //**************************************************************************
    procedure Remove(ItemIndex: Integer); overload;
    procedure Remove(Item: TKAZipEntriesEntry); overload;
    procedure Remove(FileName: string); overload;
    procedure RemoveFiles(List: TList);
    procedure RemoveSelected;
    //**************************************************************************
    procedure Select(WildCard: string);
    procedure SelectAll;
    procedure DeSelectAll;
    procedure InvertSelection;
    //**************************************************************************
    procedure Rename(Item: TKAZipEntriesEntry; NewFileName: string); overload;
    procedure Rename(ItemIndex: Integer; NewFileName: string); overload;
    procedure Rename(FileName: string; NewFileName: string); overload;
    procedure CreateFolder(FolderName: string; FolderDate: TDateTime);
    procedure RenameFolder(FolderName: string; NewFolderName: string);
    procedure RenameMultiple(Names: TStringList; NewNames: TStringList);
    //**************************************************************************
    procedure ExtractToFile(Item: TKAZipEntriesEntry; FileName: string); overload;
    procedure ExtractToFile(ItemIndex: Integer; FileName: string); overload;
    procedure ExtractToFile(FileName, DestinationFileName: string); overload;
    procedure ExtractToStream(Item: TKAZipEntriesEntry; TargetStream: TStream);
    procedure ExtractAll(TargetDirectory: string);
    procedure ExtractSelected(TargetDirectory: string);
    //**************************************************************************
    property Entries: TKAZipEntries read FZipHeader;
    property HasBadEntries: Boolean read FHasBadEntries;
    property ZipStream: TStream read FZipStream;
  published
    { Published declarations }
    property FileName: string read FFileName write SetFileName;
    property IsZipFile: Boolean read FIsZipFile write SetIsZipFile;
    property SaveMethod: TZipSaveMethod read FZipSaveMethod write SetZipSaveMethod;
    property StoreRelativePath: Boolean read FStoreRelativePath write FStoreRelativePath;
    property StoreFolders: Boolean read FStoreFolders write SetStoreFolders;
    property CompressionType: TZipCompressionType read FZipCompressionType write SetZipCompressionType;
    property Comment: TStrings read GetComment write SetComment;
    property FileNames: TStrings read GetFileNames write SetFileNames;
    property UseTempFiles: Boolean read FUseTempFiles write SetUseTempFiles;
    property OverwriteAction: TOverwriteAction read FOverwriteAction write SetOverwriteAction;
    property ComponentVersion: string read FComponentVersion write SetComponentVersion;
    property ReadOnly: Boolean read FReadOnly write SetReadOnly;
    property ApplyAtributes: Boolean read FApplyAttributes write SetApplyAtributes;
    property OnDecompressFile: TOnDecompressFile read FOnDecompressFile write FOnDecompressFile;
    property OnCompressFile: TOnCompressFile read FOnCompressFile write FOnCompressFile;
    property OnZipChange: TOnZipChange read FOnZipChange write FOnZipChange;
    property OnZipOpen: TOnZipOpen read FOnZipOpen write FOnZipOpen;
    property OnAddItem: TOnAddItem read FOnAddItem write SetOnAddItem;
    property OnRebuildZip: TOnRebuildZip read FOnRebuildZip write SetOnRebuildZip;
    property OnRemoveItems: TOnRemoveItems read FOnRemoveItems write SetOnRemoveItems;
    property OnOverwriteFile: TOnOverwriteFile read FOnOverwriteFile write SetOnOverwriteFile;
    property Active: Boolean read FIsZipFile write SetActive;
  end;

type
  TCRC32Stream = class(TStream)
  private
    FTargetStream: TStream;
    FTargetStreamOwnership: TStreamOwnership;
    FCRC32: LongWord;
    FTotalBytes: Int64;
    class function CalcCRC32(const Data: AnsiString): LongWord;
    function GetCRC32: LongWord;
  public
    constructor Create(TargetStream: TStream; StreamOwnership: TStreamOwnership=soReference);
    destructor Destroy; override;

    class procedure SelfTest;

    function Read(var Buffer; Count: Longint): Longint; override;
    function Write(const Buffer; Count: Longint): Longint; override;
    function Seek(Offset: Longint; Origin: Word): Longint; override;

    property CRC32: LongWord read GetCRC32;
    property TotalBytesWritten: Int64 read FTotalBytes;
  end;

  TKAZIPStream = class(TStream)
  private
    FParent: TKAZip;
    FTargetStream: TStream;
    FEntry: TKAZipEntriesEntry;
    FStartPosition: Int64; //used to calculate size of compressed data

    //internal stuff
    FCompressor: TZCompressionStream;
    FCrc32Stream: TCRC32Stream;
  protected
    constructor Create(TargetStream: TStream; Entry: TKAZipEntriesEntry; ZipCompressionType: TZipCompressionType; ParentZip: TKAZip);
    procedure FinalizeEntryToZip(const uncompressedLength: Int64; const crc32: Longword);
  public
    destructor Destroy; override;

    function Read(var Buffer; Count: Longint): Longint; override;
    function Write(const Buffer; Count: Longint): Longint; override;
    function Seek(Offset: Longint; Origin: Word): Longint; override;
  end;

(* for internal usage
function ReadBA(MS: TStream; Sz, Poz: Integer): TBytes;
*)
function Adler32(adler: uLong; buf: pByte; len: uInt): uLong;
function CalcCRC32(const UncompressedData: string): Cardinal;
function CalculateCRCFromStream(Stream: TStream): Cardinal;
function RemoveRootName(const FileName, RootName: string): string;


procedure Register;
function ToZipName(FileName: string): string;
function ToDosName(FileName: string): string;

implementation

uses
  {$IF CompilerVersion >= 24} // >= XE3
    Vcl.FileCtrl;
  {$ELSE}
     FileCtrl;
  {$IFEND}

  {
    We use Zip compression method 8 (Deflate).
    We don't have an actual Deflate algorithm (RFC 1951) handy.
    But we can use zlib. Zlib supports many compression algorithms (RFC 1950).
    One of those is Deflate. We we'll use zlib to compress/decompress.
    The only issue is that zlib added a 2 byte head, and 4 byte Adler checksum trailer

      CMF (1 byte)
      FLG (1 byte)
      ...DEFLATE compressed data...
      ADLER32 (4 bytes)

        CMF:  low 4 bits: Compression Method (CM):  8 = Deflate
            high 4 bits: Compression Info (CINFO): Log(b=2)(WindowSize)-8     e.g. 7 ==> 32k window size  (7+8)=15, 2^15=32768
        FLG:  low 4 bits: Check bits for CMF|FLG  (CMF*256 + FLG mod 31 = 0)
               5 bit:  Dictionary present
              78 bits: Compression level

            0 - compressor used fastest algorithm
            1 - compressor used fast algorithm
            2 - compressor used default algorithm
            3 - compressor used maximum compression, slowest algorithm

    So we have to strip off these leading 2 bytes and trailing 4 bytes in order to extract the original DEFLATE data.
    This makes streaming compression hard
  }

const
  ZL_CompressionMethod_Deflate =             $8; //ZLIB Compression Method code for DEFLATE algorithm
  ZL_Deflate_CompressionInfo_DefaultWindowSize =  $7; //ZLIB. When using DEFLATE describes the window side (7=32k)
  ZL_PRESET_DICT = $20;

  ZL_FASTEST_COMPRESSION = $0;    //0 = The file is stored (no compression)
  ZL_FAST_COMPRESSION = $1;      //1 = The file is Shrunk
  ZL_DEFAULT_COMPRESSION = $2;    //2 = The file is Reduced with compression factor 1
  ZL_MAXIMUM_COMPRESSION = $3;    //3 = The file is Reduced with compression factor 2

  ZL_FCHECK_MASK = $1F;
  ZL_CINFO_MASK = $F0; { mask out leftmost 4 bits }
  ZL_FLEVEL_MASK = $C0; { mask out leftmost 2 bits }
  ZL_CM_MASK = $0F; { mask out rightmost 4 bits }

  SIG_CentralFile: TZipSignature =            'PK'#1#2; // $02014B50
  SIG_LocalHeader: TZipSignature =            'PK'#3#4; // $04034B50
  SIG_EndOfCentralDirectory: TZipSignature =  'PK'#5#6; // $06054B50
  SIG_ExtraDataRecord: TZipSignature =        'PK'#6#8; // $08064B50
  ZL_MULTIPLE_DISK_SIG: TZipSignature =       'PK'#7#8; // $08074B50
  ZL_DATA_DESCRIPT_SIG: TZipSignature =       'PK'#7#8; // $08074B50
type
  TZLibStreamHeader = packed record
    CMF: Byte;
    FLG: Byte;
  end;

const
  CRCTable: array[0..255] of Cardinal = (
    $00000000, $77073096, $EE0E612C, $990951BA, $076DC419, $706AF48F, $E963A535,
    $9E6495A3, $0EDB8832, $79DCB8A4, $E0D5E91E, $97D2D988, $09B64C2B, $7EB17CBD,
    $E7B82D07, $90BF1D91, $1DB71064, $6AB020F2, $F3B97148, $84BE41DE, $1ADAD47D,
    $6DDDE4EB, $F4D4B551, $83D385C7, $136C9856, $646BA8C0, $FD62F97A, $8A65C9EC,
    $14015C4F, $63066CD9, $FA0F3D63, $8D080DF5, $3B6E20C8, $4C69105E, $D56041E4,
    $A2677172, $3C03E4D1, $4B04D447, $D20D85FD, $A50AB56B, $35B5A8FA, $42B2986C,
    $DBBBC9D6, $ACBCF940, $32D86CE3, $45DF5C75, $DCD60DCF, $ABD13D59, $26D930AC,
    $51DE003A, $C8D75180, $BFD06116, $21B4F4B5, $56B3C423, $CFBA9599, $B8BDA50F,
    $2802B89E, $5F058808, $C60CD9B2, $B10BE924, $2F6F7C87, $58684C11, $C1611DAB,
    $B6662D3D, $76DC4190, $01DB7106, $98D220BC, $EFD5102A, $71B18589, $06B6B51F,
    $9FBFE4A5, $E8B8D433, $7807C9A2, $0F00F934, $9609A88E, $E10E9818, $7F6A0DBB,
    $086D3D2D, $91646C97, $E6635C01, $6B6B51F4, $1C6C6162, $856530D8, $F262004E,
    $6C0695ED, $1B01A57B, $8208F4C1, $F50FC457, $65B0D9C6, $12B7E950, $8BBEB8EA,
    $FCB9887C, $62DD1DDF, $15DA2D49, $8CD37CF3, $FBD44C65, $4DB26158, $3AB551CE,
    $A3BC0074, $D4BB30E2, $4ADFA541, $3DD895D7, $A4D1C46D, $D3D6F4FB, $4369E96A,
    $346ED9FC, $AD678846, $DA60B8D0, $44042D73, $33031DE5, $AA0A4C5F, $DD0D7CC9,
    $5005713C, $270241AA, $BE0B1010, $C90C2086, $5768B525, $206F85B3, $B966D409,
    $CE61E49F, $5EDEF90E, $29D9C998, $B0D09822, $C7D7A8B4, $59B33D17, $2EB40D81,
    $B7BD5C3B, $C0BA6CAD, $EDB88320, $9ABFB3B6, $03B6E20C, $74B1D29A, $EAD54739,
    $9DD277AF, $04DB2615, $73DC1683, $E3630B12, $94643B84, $0D6D6A3E, $7A6A5AA8,
    $E40ECF0B, $9309FF9D, $0A00AE27, $7D079EB1, $F00F9344, $8708A3D2, $1E01F268,
    $6906C2FE, $F762575D, $806567CB, $196C3671, $6E6B06E7, $FED41B76, $89D32BE0,
    $10DA7A5A, $67DD4ACC, $F9B9DF6F, $8EBEEFF9, $17B7BE43, $60B08ED5, $D6D6A3E8,
    $A1D1937E, $38D8C2C4, $4FDFF252, $D1BB67F1, $A6BC5767, $3FB506DD, $48B2364B,
    $D80D2BDA, $AF0A1B4C, $36034AF6, $41047A60, $DF60EFC3, $A867DF55, $316E8EEF,
    $4669BE79, $CB61B38C, $BC66831A, $256FD2A0, $5268E236, $CC0C7795, $BB0B4703,
    $220216B9, $5505262F, $C5BA3BBE, $B2BD0B28, $2BB45A92, $5CB36A04, $C2D7FFA7,
    $B5D0CF31, $2CD99E8B, $5BDEAE1D, $9B64C2B0, $EC63F226, $756AA39C, $026D930A,
    $9C0906A9, $EB0E363F, $72076785, $05005713, $95BF4A82, $E2B87A14, $7BB12BAE,
    $0CB61B38, $92D28E9B, $E5D5BE0D, $7CDCEFB7, $0BDBDF21, $86D3D2D4, $F1D4E242,
    $68DDB3F8, $1FDA836E, $81BE16CD, $F6B9265B, $6FB077E1, $18B74777, $88085AE6,
    $FF0F6A70, $66063BCA, $11010B5C, $8F659EFF, $F862AE69, $616BFFD3, $166CCF45,
    $A00AE278, $D70DD2EE, $4E048354, $3903B3C2, $A7672661, $D06016F7, $4969474D,
    $3E6E77DB, $AED16A4A, $D9D65ADC, $40DF0B66, $37D83BF0, $A9BCAE53, $DEBB9EC5,
    $47B2CF7F, $30B5FFE9, $BDBDF21C, $CABAC28A, $53B39330, $24B4A3A6, $BAD03605,
    $CDD70693, $54DE5729, $23D967BF, $B3667A2E, $C4614AB8, $5D681B02, $2A6F2B94,
    $B40BBE37, $C30C8EA1, $5A05DF1B, $2D02EF8D);

const
  Default_TCentralDirectoryFile: TCentralDirectoryFile = ();


procedure Register;
begin
  RegisterComponents('KA', [TKAZip]);
end;

function ReadBA(MS: TStream; Sz, Poz: Integer): TBytes;
begin
  SetLength(Result, SZ);
  MS.Position := Poz;
  MS.Read(Result[0], SZ);
end;

function Adler32(adler: uLong; buf: pByte; len: uInt): uLong;
const
  BASE = uLong(65521);
  NMAX = 3854;
var
  s1, s2: uLong;
  k: Integer;
begin
  s1 := adler and $FFFF;
  s2 := (adler shr 16) and $FFFF;

  if not Assigned(buf) then
  begin
    adler32 := uLong(1);
    exit;
  end;

  while (len > 0) do
  begin
    if len < NMAX then
      k := len
    else
      k := NMAX;
    Dec(len, k);
    while (k > 0) do
    begin
      Inc(s1, buf^);
      Inc(s2, s1);
      Inc(buf);
      Dec(k);
    end;
    s1 := s1 mod BASE;
    s2 := s2 mod BASE;
  end;
  adler32 := (s2 shl 16) or s1;
end;

function CalcCRC32(const UncompressedData: string): Cardinal;
var
  X: Integer;
begin
  Result := $FFFFFFFF;
  for X := 0 to Length(UncompressedData) - 1 do
  begin
    Result := (Result shr 8) xor (CRCTable[Byte(Result) xor Ord(UncompressedData[X + 1])]);
  end;
  Result := Result xor $FFFFFFFF;
end;

function CalculateCRCFromStream(Stream: TStream): Cardinal;
var
  Buffer: array[1..8192] of Byte;
  I, ReadCount: Integer;
  TempResult: Longword;
begin
  TempResult := $FFFFFFFF;
  while (Stream.Position <> Stream.Size) do
  begin
    ReadCount := Stream.Read(Buffer, SizeOf(Buffer));
    for I := 1 to ReadCount do
      TempResult := ((TempResult shr 8) and $FFFFFF) xor CRCTable[(TempResult xor Longword(Buffer[I])) and $FF];
  end;
  Result := not TempResult;
end;

function RemoveRootName(const FileName, RootName: string): string;
var
  P: Integer;
  S: string;
begin
  Result := FileName;
  P := Pos(AnsiLowerCase(RootName), AnsiLowerCase(FileName));
  if P = 1 then
  begin
    System.Delete(Result, 1, Length(RootName));
    S := Result;
    if (Length(S) > 0) and (S[1] = '\') then
    begin
      System.Delete(S, 1, 1);
      Result := S;
    end;
  end;
end;

function ToZipName(FileName: string): string;
var
  P: Integer;
begin
  Result := FileName;
  Result := StringReplace(Result, '\', '/', [rfReplaceAll]);
  P := Pos(':/', Result);
  if P > 0 then
  begin
    System.Delete(Result, 1, P + 1);
  end;
  P := Pos('//', Result);
  if P > 0 then
  begin
    System.Delete(Result, 1, P + 1);
    P := Pos('/', Result);
    if P > 0 then
    begin
      System.Delete(Result, 1, P);
      P := Pos('/', Result);
      if P > 0 then
        System.Delete(Result, 1, P);
    end;
  end;
end;

function ToDosName(FileName: string): string;
var
  P: Integer;
begin
  Result := FileName;
  Result := StringReplace(Result, '\', '/', [rfReplaceAll]);
  P := Pos(':/', Result);
  if P > 0 then
  begin
    System.Delete(Result, 1, P + 1);
  end;
  P := Pos('//', Result);
  if P > 0 then
  begin
    System.Delete(Result, 1, P + 1);
    P := Pos('/', Result);
    if P > 0 then
    begin
      System.Delete(Result, 1, P);
      P := Pos('/', Result);
      if P > 0 then
        System.Delete(Result, 1, P);
    end;
  end;
  Result := StringReplace(Result, '/', '\', [rfReplaceAll]);
end;

function SeekForSignature(MS: TStream; ASig: TZipSignature; var APos: Int64): Boolean;
var
  MaxPos: Int64;
  TmpSig: TZipSignature;
begin
  MaxPos := MS.Size - SizeOf(TmpSig);
  Result := False;
  while APos <= MaxPos do
  begin
    MS.Position := APos;
    MS.Read(TmpSig, SizeOf(TmpSig));
    if TmpSig[0] = ASig[0] then
    begin
      if TmpSig = ASig then
      begin
        Result := True;
        Exit;
      end
      else
      begin
        Inc(APos, 4);
      end;
    end
    else
    begin
      Inc(APos);
    end;
  end;
end;

{ TKAZipEntriesEntry }

constructor TKAZipEntriesEntry.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FParent := TKAZipEntries(Collection);
  FSelected := False;
  FVersionMadeBy := 20;
  FVersionNeededToExtract := 20;
  FGeneralPurposeBitFlag := 0;
  //FDate := Now();
end;

destructor TKAZipEntriesEntry.Destroy;
begin

  inherited Destroy;
end;

procedure TKAZipEntriesEntry.UpdateProperties();
begin
  FCompressionType := ctUnknown;
  if (FCompressionMethod = ZipCompressionMethod_Deflate)
  or (FCompressionMethod = ZipCompressionMethod_Deflate64) then
  begin
    case FGeneralPurposeBitFlag and 6 of
      0: FCompressionType := ctNormal;
      2: FCompressionType := ctMaximum;
      4: FCompressionType := ctFast;
      6: FCompressionType := ctSuperFast
    end;
  end;

  if (FGeneralPurposeBitFlag and 1) > 0 then
    FIsEncrypted := True
  else
    FIsEncrypted := False;

  FIsFolder := (FExternalFileAttributes and faDirectory) > 0;
end;

function TKAZipEntriesEntry.ParseFromCentralDirectory(MS: TStream): Boolean;
var
  s: AnsiString;
  CDFile: TCentralDirectoryFile;
begin
  Result := False;
  try
    //FillCh ar(CDFile, SizeOf(TCentralDirectoryFile), 0);
    // 20140320: Filling a record containing managed types will cause a leak
    CDFile := Default_TCentralDirectoryFile;
    MS.Read(CDFile, SizeOf(CDFile));

    if CDFile.CentralFileHeaderSignature <> SIG_CentralFile then
      Exit;

    FVersionMadeBy := CDFile.VersionMadeBy;
    FVersionNeededToExtract := CDFile.VersionNeededToExtract;
    FGeneralPurposeBitFlag := CDFile.GeneralPurposeBitFlag;
    FCompressionMethod := CDFile.CompressionMethod;
    FDate := FileDateToDateTime(CDFile.LastModFileTimeDate);
    FCRC32 := CDFile.Crc32;
    FSizeCompressed := CDFile.CompressedSize;
    FSizeUncompressed := CDFile.UncompressedSize;
    FDiskNumber := CDFile.DiskNumberStart;
    FInternalFileAttributes := CDFile.InternalFileAttributes;
    FExternalFileAttributes := CDFile.ExternalFileAttributes;
    FLocalFileHeaderOffset := CDFile.RelativeOffsetOfLocalHeader;

    if CDFile.FilenameLength > 0 then
    begin
      SetLength(s, CDFile.FilenameLength);
      MS.Read(s[1], CDFile.FilenameLength);
      FFileName := s;
    end;
    if CDFile.ExtraFieldLength > 0 then
    begin
      SetLength(s, CDFile.ExtraFieldLength);
      MS.Read(s[1], CDFile.ExtraFieldLength);
      FExtraField := s;
    end;
    if CDFile.FileCommentLength > 0 then
    begin
      SetLength(s, CDFile.FileCommentLength);
      MS.Read(s[1], CDFile.FileCommentLength);
      FFileComment := s;
    end;

    UpdateProperties();
    Result := True;
  except
    Exit;
  end;
end;

function TKAZipEntriesEntry.ParseFromLocalHeader(MS: TStream): Boolean;
var
  s: AnsiString;
  HeaderPos, NextPos: Int64;
  LFH: TLocalFileHeader;
  DataDescriptor: TDataDescriptor;
begin
  Result := False;
  try
    // todo: check signature
    HeaderPos := MS.Position;
    MS.Read(LFH, SizeOf(LFH));
    FVersionNeededToExtract := LFH.VersionNeededToExtract;
    FGeneralPurposeBitFlag := LFH.GeneralPurposeBitFlag;
    FCompressionMethod := LFH.CompressionMethod;
    FDate := FileDateToDateTime(LFH.LastModFileTimeDate);
    if LFH.Crc32 <> 0 then
      FCRC32 := LFH.Crc32;
    if LFH.CompressedSize <> 0 then
      FSizeCompressed := LFH.CompressedSize;
    if LFH.UncompressedSize <> 0 then
      FSizeUncompressed := LFH.UncompressedSize;
    if LFH.FilenameLength > 0 then
    begin
      SetLength(s, LFH.FilenameLength);
      MS.Read(s[1], LFH.FilenameLength);
      FFileName := s;

      if FFileName[LFH.FilenameLength] = '/' then
        FExternalFileAttributes := faDirectory;
    end;
    if LFH.ExtraFieldLength > 0 then
    begin
      SetLength(s, LFH.ExtraFieldLength);
      MS.Read(s[1], LFH.ExtraFieldLength);
      FExtraField := s;
    end;
    FCompressedDataPos := MS.Position;

    if ((LFH.GeneralPurposeBitFlag and (1 shl 3)) > 0) and (FCrc32 = 0) then
    begin
      NextPos := HeaderPos + SizeOf(LFH);
      if SeekForSignature(MS, SIG_LocalHeader, NextPos) then
      begin
        MS.Position := NextPos - SizeOf(TDataDescriptor);
        { TODO: check optional signature }
        MS.Read(DataDescriptor, SizeOf(TDataDescriptor));
        FCrc32 := DataDescriptor.Crc32;
        FSizeCompressed := DataDescriptor.CompressedSize;
        FSizeUncompressed := DataDescriptor.UncompressedSize;
      end;
    end;

    FLocalFileHeaderOffset := HeaderPos;

    UpdateProperties();

    Result := True;
  except
    Exit;
  end;
end;

procedure TKAZipEntriesEntry.FillCentralDirectory(out CDFile: TCentralDirectoryFile);
begin
  UpdateProperties();
  CDFile.CentralFileHeaderSignature := SIG_CentralFile;
  CDFile.VersionMadeBy := FVersionMadeBy;
  CDFile.VersionNeededToExtract := FVersionNeededToExtract;
  CDFile.GeneralPurposeBitFlag := FGeneralPurposeBitFlag;
  CDFile.CompressionMethod := FCompressionMethod;
  CDFile.LastModFileTimeDate := LongWord(DateTimeToFileDate(FDate));
  CDFile.Crc32 := FCRC32;
  CDFile.CompressedSize := FSizeCompressed;
  CDFile.UncompressedSize := FSizeUncompressed;
  CDFile.FilenameLength := Length(FileName);
  CDFile.ExtraFieldLength := Length(ExtraField);
  CDFile.FileCommentLength := Length(Comment);
  CDFile.DiskNumberStart := FDiskNumber;
  CDFile.InternalFileAttributes := FInternalFileAttributes;
  CDFile.ExternalFileAttributes := FExternalFileAttributes;
  CDFile.RelativeOffsetOfLocalHeader := FLocalFileHeaderOffset;
end;

procedure TKAZipEntriesEntry.FillLocalHeader(out LFH: TLocalFileHeader);
begin
  UpdateProperties();
  LFH.LocalFileHeaderSignature := SIG_LocalHeader; //'PK'#03#04
  LFH.VersionNeededToExtract := FVersionNeededToExtract;
  LFH.GeneralPurposeBitFlag := FGeneralPurposeBitFlag;
  LFH.CompressionMethod := FCompressionMethod;
  LFH.LastModFileTimeDate := LongWord(DateTimeToFileDate(FDate));
  LFH.Crc32 := FCRC32;
  LFH.CompressedSize := FSizeCompressed;
  LFH.UncompressedSize := FSizeUncompressed;
  LFH.FilenameLength := Length(FFileName);
  LFH.ExtraFieldLength := Length(FExtraField);
end;

procedure TKAZipEntriesEntry.ExtractToFile(FileName: string);
begin
  FParent.ExtractToFile(Self, FileName);
end;

procedure TKAZipEntriesEntry.ExtractToStream(Stream: TStream);
begin
  FParent.ExtractToStream(Self, Stream);
end;

procedure TKAZipEntriesEntry.SaveToFile(FileName: string);
begin
  ExtractToFile(FileName);
end;

procedure TKAZipEntriesEntry.SaveToStream(Stream: TStream);
begin
  ExtractToStream(Stream);
end;

function TKAZipEntriesEntry.GetCompressedData(Stream: TStream): Integer;
var
  ZLHeader: TZLibStreamHeader;
  ZLH: Word;
  compress: Byte;
begin
  Result := 0;

  if (CompressionMethod = ZipCompressionMethod_Deflate) then
  begin
    //Reconstruct the 2-byte ZLib compression header (CMF & FLG)
    //We have to do this because the data is DEFLATE compressed, but we only have a zlib library
    //So have to reverse engineer what zlib originally would have put there

    ZLHeader.CMF := ZL_CompressionMethod_Deflate; { Deflate }
    ZLHeader.CMF := ZLHeader.CMF or (ZL_Deflate_CompressionInfo_DefaultWindowSize shl 4); { 32k Window size }
    compress := ZL_DEFAULT_COMPRESSION;
    case BitFlag and 6 of
      0: Compress := ZL_DEFAULT_COMPRESSION;
      2: Compress := ZL_MAXIMUM_COMPRESSION;
      4: Compress := ZL_FAST_COMPRESSION;
      6: Compress := ZL_FASTEST_COMPRESSION;
    end;
    ZLHeader.FLG := ZLHeader.FLG or (Compress shl 6);
    ZLHeader.FLG := ZLHeader.FLG and not ZL_PRESET_DICT; { no preset dictionary}
    ZLHeader.FLG := ZLHeader.FLG and not ZL_FCHECK_MASK;
    ZLH := (ZLHeader.CMF * 256) + ZLHeader.FLG;
    Inc(ZLHeader.FLG, 31 - (ZLH mod 31));

    //Now write the zlib preamble
    Result := Result + Stream.Write(ZLHeader, SizeOf(ZLHeader));
  end;

  if SizeCompressed > 0 then
  begin
    if FCompressedDataPos = 0 then
      FCompressedDataPos := LocalOffset + GetLocalHeadSize();

    FParent.FParent.ZipStream.Position := FCompressedDataPos;
    Result := Result + Stream.CopyFrom(FParent.FParent.ZipStream, SizeCompressed);
  end;
end;

function TKAZipEntriesEntry.GetCompressedData: AnsiString;
var
  ResultSize: Integer;
  ZLHeader: TZLibStreamHeader;
  ZLH: Word;
  Compress: Byte;
begin
  Result := '';

  case CompressionMethod of
    ZipCompressionMethod_Store,
    ZipCompressionMethod_Deflate:
    begin
      if (CompressionMethod = ZipCompressionMethod_Deflate) then
      begin
        ZLHeader.CMF := ZL_CompressionMethod_Deflate; { Deflate }
        ZLHeader.CMF := ZLHeader.CMF or (ZL_Deflate_CompressionInfo_DefaultWindowSize shl 4); { 32k Window size }
        Compress := ZL_DEFAULT_COMPRESSION;
        case BitFlag and 6 of
          0: Compress := ZL_DEFAULT_COMPRESSION;
          2: Compress := ZL_MAXIMUM_COMPRESSION;
          4: Compress := ZL_FAST_COMPRESSION;
          6: Compress := ZL_FASTEST_COMPRESSION;
        end;
        ZLHeader.FLG := ZLHeader.FLG or (Compress shl 6);
        ZLHeader.FLG := ZLHeader.FLG and not ZL_PRESET_DICT; { no preset dictionary}
        ZLHeader.FLG := ZLHeader.FLG and not ZL_FCHECK_MASK;
        ZLH := (ZLHeader.CMF * 256) + ZLHeader.FLG;
        Inc(ZLHeader.FLG, 31 - (ZLH mod 31));
        SetLength(Result, SizeOf(ZLHeader));
        SetString(Result, PChar(@ZLHeader), SizeOf(ZLHeader));
      end;
      if SizeCompressed > 0 then
      begin
        if FCompressedDataPos = 0 then
          FCompressedDataPos := LocalOffset + GetLocalHeadSize();
        ResultSize := Length(Result);
        SetLength(Result, ResultSize + SizeCompressed);
        FParent.FParent.ZipStream.Position := CompressedDataPos;
        FParent.FParent.ZipStream.Read(Result[ResultSize+1], CompressedDataPos);
      end;
    end;
  end;
end;

procedure TKAZipEntriesEntry.SetSelected(const Value: Boolean);
begin
  FSelected := Value;
end;

function TKAZipEntriesEntry.GetLocalHeadSize: LongWord;
begin
  Result := SizeOf(TLocalFileHeader) + Length(FFileName) + Length(FExtraField);
end;

function TKAZipEntriesEntry.GetLocalEntrySize: LongWord;
begin
  if FLocalFileEntrySize = 0 then
  begin
    FLocalFileEntrySize := SizeOf(TLocalFileHeader) + FSizeCompressed
            + Length(FFileName) + Length(FExtraField);
    if (FGeneralPurposeBitFlag and (1 shl 3)) > 0 then
    begin
      FLocalFileEntrySize := FLocalFileEntrySize + SizeOf(TDataDescriptor);
    end;
  end;
  Result := FLocalFileEntrySize;
end;

function TKAZipEntriesEntry.GetCentralEntrySize: LongWord;
begin
  if FCentralFileEntrySize = 0 then
    FCentralFileEntrySize := SizeOf(TCentralDirectoryFile)
      + Length(FFileName) + Length(FExtraField) + Length(FFileComment);
  Result := FCentralFileEntrySize;
end;

function TKAZipEntriesEntry.Test: Boolean;
var
  FS: TFileStream;
  MS: TMemoryStream;
  FN: string;
begin
  Result := True;
  try
    if not FIsEncrypted then
    begin
      if FParent.FParent.FUseTempFiles then
      begin
        FN := FParent.FParent.GetDelphiTempFileName;
        FS := TFileStream.Create(FN, fmOpenReadWrite or FmCreate);
        try
          ExtractToStream(FS);
          FS.Position := 0;
          Result := CalculateCRCFromStream(FS) = CRC32;
        finally
          FS.Free;
          DeleteFile(FN);
        end;
      end
      else
      begin
        MS := TMemoryStream.Create;
        try
          ExtractToStream(MS);
          MS.Position := 0;
          Result := CalculateCRCFromStream(MS) = CRC32;
        finally
          MS.Free;
        end;
      end;
    end;
  except
    Result := False;
  end;
end;

procedure TKAZipEntriesEntry.SetComment(const Value: AnsiString);
begin
  FFileComment := Value;
  FCentralFileEntrySize := 0;
  //FParent.Rebuild;
  if not FParent.FParent.FBatchMode then
  begin
    FParent.FParent.DoChange(FParent, ZIP_ITEM_COMMENT);
  end;
end;

procedure TKAZipEntriesEntry.SetFileName(const Value: AnsiString);
var
  FN: string;
begin
  FN := ToZipName(Value);
  if FParent.IndexOf(FN) > -1 then
    raise Exception.Create('File with same name already exists in Archive!');
  FFileName := FN;
  FCentralFileEntrySize := 0;
  if not FParent.FParent.FBatchMode then
  begin
    //FParent.Rebuild;
    FParent.FParent.DoChange(FParent, ZIP_ITEM_NAME);
  end;
end;

{ TKAZipEntries }

constructor TKAZipEntries.Create(AOwner: TKAZip);
begin
  inherited Create(TKAZipEntriesEntry);
  FParent := AOwner;
end;

constructor TKAZipEntries.Create(AOwner: TKAZip; MS: TStream);
begin
  inherited Create(TKAZipEntriesEntry);
  FParent := AOwner;
  ParseZip(MS);
end;

destructor TKAZipEntries.Destroy;
begin

  inherited Destroy;
end;

procedure TKAZipEntries.ParseZip(MS: TStream);
begin
  FParent.ParseZip(MS);
end;

procedure TKAZipEntries.SortList(List: TList);
var
  X: Integer;
  I1: Cardinal;
  I2: Cardinal;
  NoChange: Boolean;
begin
  if List.Count = 1 then
    Exit;
  repeat
    NoChange := True;
    for X := 0 to List.Count - 2 do
    begin
      I1 := Integer(List.Items[X]);
      I2 := Integer(List.Items[X + 1]);
      if I1 > I2 then
      begin
        List.Exchange(X, X + 1);
        NoChange := False;
      end;
    end;
  until NoChange;
end;

function TKAZipEntries.FileTime2DateTime(FileTime: TFileTime): TDateTime;
var
  LocalFileTime: TFileTime;
  SystemTime: TSystemTime;
begin
  FileTimeToLocalFileTime(FileTime, LocalFileTime);
  FileTimeToSystemTime(LocalFileTime, SystemTime);
  Result := SystemTimeToDateTime(SystemTime);
end;

function TKAZipEntries.GetHeaderEntry(Index: Integer): TKAZipEntriesEntry;
begin
  Result := TKAZipEntriesEntry(inherited Items[Index]);
end;

procedure TKAZipEntries.SetHeaderEntry(Index: Integer; const Value: TKAZipEntriesEntry);
begin
  inherited Items[Index] := TCollectionItem(Value);
end;

procedure TKAZipEntries.Remove(ItemIndex: Integer);
begin
  FParent.Remove(ItemIndex, True);
end;

procedure TKAZipEntries.Remove(Item: TKAZipEntriesEntry);
var
  X: Integer;
begin
  for X := 0 to Count - 1 do
  begin
    if Self.Items[X] = Item then
    begin
      Remove(X);
      Exit;
    end;
  end;
end;

procedure TKAZipEntries.Remove(FileName: string);
var
  I: Integer;
begin
  I := IndexOf(FileName);
  if I <> -1 then
    Remove(I);
end;

function TKAZipEntries.IndexOf(const FileName: string): Integer;
var
  X: Integer;
  FN: string;
begin
  Result := -1;
  FN := ToZipName(FileName);
  for X := 0 to Count - 1 do
  begin
    if AnsiCompareText(FN, ToZipName(Items[X].FileName)) = 0 then
    begin
      Result := X;
      Exit;
    end;
  end;
end;

function TKAZipEntries.AddStream(FileName: string; FileAttr: Word;
  FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := FParent.AddStream(FileName, FileAttr, FileDate, Stream);
end;

function TKAZipEntries.AddStream(FileName: string; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := FParent.AddStream(FileName, faArchive, Now, Stream);
end;

function TKAZipEntries.AddFile(FileName, NewFileName: string): TKAZipEntriesEntry;
begin
  Result := FParent.AddFile(FileName, NewFileName);
end;

function TKAZipEntries.AddFile(FileName: string): TKAZipEntriesEntry;
begin
  Result := AddFile(FileName, FileName);
end;

function TKAZipEntries.AddFiles(FileNames: TStrings): Boolean;
begin
  Result := FParent.AddFiles(FileNames);
end;

function TKAZipEntries.AddFolder(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
begin
  Result := FParent.AddFolder(FolderName, RootFolder, WildCard, WithSubFolders);
end;

function TKAZipEntries.AddFilesAndFolders(FileNames: TStrings; RootFolder: string; WithSubFolders: Boolean): Boolean;
begin
  Result := FParent.AddFilesAndFolders(FileNames, RootFolder, WithSubFolders);
end;

procedure TKAZipEntries.RemoveFiles(List: TList);
begin
  FParent.RemoveFiles(List);
end;

procedure TKAZipEntries.RemoveSelected;
begin
  FParent.RemoveSelected;
end;

procedure TKAZipEntries.ExtractToStream(Item: TKAZipEntriesEntry; TargetStream: TStream);
begin
  FParent.ExtractToStream(Item, TargetStream);
end;

procedure TKAZipEntries.ExtractToFile(Item: TKAZipEntriesEntry; FileName: string);
begin
  FParent.ExtractToFile(Item, FileName);
end;

procedure TKAZipEntries.ExtractToFile(ItemIndex: Integer; FileName: string);
begin
  FParent.ExtractToFile(ItemIndex, FileName);
end;

procedure TKAZipEntries.ExtractToFile(FileName, DestinationFileName: string);
begin
  FParent.ExtractToFile(FileName, DestinationFileName);
end;

procedure TKAZipEntries.ExtractAll(TargetDirectory: string);
begin
  FParent.ExtractAll(TargetDirectory);
end;

procedure TKAZipEntries.ExtractSelected(TargetDirectory: string);
begin
  FParent.ExtractSelected(TargetDirectory);
end;

procedure TKAZipEntries.DeSelectAll;
var
  X: Integer;
begin
  for X := 0 to Count - 1 do
    Items[X].Selected := False;
end;

procedure TKAZipEntries.InvertSelection;
var
  X: Integer;
begin
  for X := 0 to Count - 1 do
    Items[X].Selected := not Items[X].Selected;
end;

procedure TKAZipEntries.SelectAll;
var
  X: Integer;
begin
  for X := 0 to Count - 1 do
    Items[X].Selected := True;
end;

procedure TKAZipEntries.Select(WildCard: string);
var
  X: Integer;
begin
  for X := 0 to Count - 1 do
  begin
    if MatchesMask(ToDosName(Items[X].FileName), WildCard) then
      Items[X].Selected := True;
  end;
end;

procedure TKAZipEntries.Rebuild;
begin
  FParent.Rebuild();
end;

procedure TKAZipEntries.Rename(Item: TKAZipEntriesEntry; NewFileName: string);
begin
  Item.FileName := NewFileName;
end;

procedure TKAZipEntries.Rename(ItemIndex: Integer; NewFileName: string);
begin
  Rename(Items[ItemIndex], NewFileName);
end;

procedure TKAZipEntries.Rename(FileName, NewFileName: string);
var
  I: Integer;
begin
  I := IndexOf(FileName);
  Rename(I, NewFileName);
end;

procedure TKAZipEntries.CreateFolder(FolderName: string; FolderDate: TDateTime);
begin
  FParent.CreateFolder(FolderName, FolderDate);
end;

procedure TKAZipEntries.RenameFolder(FolderName: string; NewFolderName: string);
begin
  FParent.RenameFolder(FolderName, NewFolderName);
end;

procedure TKAZipEntries.RenameMultiple(Names: TStringList; NewNames: TStringList);
begin
  FParent.RenameMultiple(Names, NewNames);
end;

function TKAZipEntries.AddEntryThroughStream(FileName: string; FileDate: TDateTime; FileAttr: Word): TStream;
begin
  Result := FParent.AddEntryThroughStream(FileName, FileDate, FileAttr);
end;

{ TKAZip }

constructor TKAZip.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FZipStream := nil;
  FOnDecompressFile := nil;
  FOnCompressFile := nil;
  FOnZipChange := nil;
  FOnZipOpen := nil;
  FOnAddItem := nil;
  FOnOverwriteFile := nil;
  FComponentVersion := '2.0';
  FBatchMode := False;
  FFileNames := TStringList.Create;
  FZipHeader := TKAZipEntries.Create(Self);
  FZipComment := TStringList.Create;
  FIsZipFile := False;
  FFileName := '';
  FCurrentDFS := 0;
  FExternalStream := False;
  FIsDirty := True;
  FHasBadEntries := False;
  FReadOnly := False;
  FLocalHeaderNumFiles := 0;

  FApplyAttributes := True;
  FOverwriteAction := oaSkip;
  FZipSaveMethod := FastSave;
  FUseTempFiles := False;
  FStoreRelativePath := True;
  FStoreFolders := True;
  FZipCompressionType := ctMaximum;
end;

destructor TKAZip.Destroy;
begin
  if Assigned(FZipStream) and (not FExternalStream) then
    FZipStream.Free;
  FZipHeader.Free;
  FZipComment.Free;
  FFileNames.Free;
  inherited Destroy;
end;

procedure TKAZip.DoChange(Sender: TObject; const ChangeType: Integer);
begin
  if Assigned(FOnZipChange) then
    FOnZipChange(Self, ChangeType);
end;

function TKAZip.FindCentralDirectory(MS: TStream): Boolean;
var
  SeekStart: Integer;
  Poz: Integer;
  BR: Integer;
  Sig: TZipSignature;
begin
  Result := False;
  if MS.Size < 22 then
    Exit;
  if MS.Size < 256 then
    SeekStart := MS.Size
  else
    SeekStart := 256;
  Poz := MS.Size - 22;
  BR := SeekStart;
  repeat
    MS.Position := Poz;
    MS.Read(Sig, SizeOf(Sig));
    if Sig[0] = SIG_EndOfCentralDirectory[0] then
    begin
      if Sig = SIG_EndOfCentralDirectory then
      begin
        MS.Position := Poz;
        FEndOfCentralDirPos := MS.Position;
        MS.Read(FEndOfCentralDir, SizeOf(FEndOfCentralDir));
        FZipCommentPos := MS.Position;
        FZipComment.Clear;
        Result := True;
      end
      else
      begin
        Dec(Poz, 4);
        Dec(BR, 4);
      end;
    end
    else
    begin
      Dec(Poz);
      Dec(BR)
    end;
    if BR < 0 then
    begin
      case SeekStart of
        256:
        begin
          SeekStart := 1024;
          Poz := MS.Size - (256 + 22);
          BR := SeekStart;
        end;
        1024:
        begin
          SeekStart := 65536;
          Poz := MS.Size - (1024 + 22);
          BR := SeekStart;
        end;
        65536:
        begin
          SeekStart := -1;
        end;
      end;
    end;
    if BR < 0 then
      SeekStart := -1;
    if MS.Size < SeekStart then
      SeekStart := -1;
  until (Result) or (SeekStart = -1);
end;

function TKAZip.ParseCentralHeaders(MS: TStream): Boolean;
var
  X: Integer;
  Entry: TKAZipEntriesEntry;
begin
  Result := False;
  try
    MS.Position := FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
    for X := 0 to FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk - 1 do
    begin
      Entry := TKAZipEntriesEntry.Create(Entries);
      if not Entry.ParseFromCentralDirectory(MS) then
      begin
        Entry.Free();
        Break;
      end;
      if Assigned(FOnZipOpen) then
        FOnZipOpen(Self, X, FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
    end;
  except
    Exit;
  end;
  Result := (Entries.Count = FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
end;

procedure TKAZip.ParseZip(MS: TStream);
begin
  FIsZipFile := False;
  Entries.Clear();
  if FindCentralDirectory(MS) then
  begin
    if ParseCentralHeaders(MS) then
    begin
      FIsZipFile := True;
      LoadLocalHeaders(MS);
    end;
  end
  else
  begin
    if ParseLocalHeaders(MS) then
    begin
      FIsZipFile := (Entries.Count > 0);
      if FIsZipFile then
        FHasBadEntries := True;
    end;
  end;
end;

(*
function TKAZip.GetLocalEntry(MS: TStream; Offset: Integer; HeaderOnly: Boolean): TLocalFile;
var
  LFH: TLocalFileHeader;
  EDR: TZipExtraDataRecord;
  CurFilePos: Int64;
  Sig: TZipSignature;
  DataDescriptor: TDataDescriptor;
  //Sig: LongWord;
begin
  //20140320 -- You cannot ZeroMe mory/FillCh ar a structure containing strings.
  //It destroys some internal housekeeping and causes a leak.
//  FillCh ar(Result, SizeOf(Result), 0);
  Result := Default_TLocalFile;


  MS.Position := Offset;
  MS.Read(Sig, SizeOf(Sig));
  if (Sig <> SIG_LocalHeader) then
    Exit;

  MS.Position := Offset;
  MS.Read(LFH, SizeOf(LFH));
  AssignLocalFileFromHeader(Result, LFH);
  if Result.FilenameLength > 0 then
  begin
    SetLength(Result.FileName, Result.FilenameLength);
    MS.Read(Result.FileName[1], Result.FilenameLength);
  end;
  if Result.ExtraFieldLength > 0 then
  begin
    SetLength(Result.ExtraField, Result.ExtraFieldLength);
    MS.Read(Result.ExtraField[1], Result.ExtraFieldLength);
  end;
  // try to read TZipExtraDataRecord
  CurFilePos := MS.Position;
  MS.Read(EDR, SizeOf(EDR));
  if EDR.Signature = SIG_ExtraDataRecord then
  begin
    if EDR.FieldLength > 0 then
    begin
      SetLength(Result.ExtraData, EDR.FieldLength);
      MS.Read(Result.ExtraData[1], EDR.FieldLength);
    end;
  end
  else
  begin
    // rewinf back
    MS.Position := CurFilePos;
  end;
  // optional data descriptor written AFTER compressed data
  {if (Result.GeneralPurposeBitFlag and (1 shl 3)) > 0 then
  begin
    MS.Read(DataDescriptor, SizeOf(TDataDescriptor));
    Result.Crc32 := DataDescriptor.Crc32;
    Result.CompressedSize := DataDescriptor.CompressedSize;
    Result.UnCompressedSize := DataDescriptor.UnCompressedSize;
  end;  }
  if not HeaderOnly then
  begin
    if Result.CompressedSize > 0 then
    begin
      SetLength(Result.CompressedData, Result.CompressedSize);
      MS.Read(Result.CompressedData[1], Result.CompressedSize);
    end;
  end;
end;
*)

procedure TKAZip.LoadLocalHeaders(MS: TStream);
var
  i: Integer;
begin
  FHasBadEntries := False;
  for i := 0 to Entries.Count - 1 do
  begin
    if Assigned(FOnZipOpen) then
      FOnZipOpen(Self, i, FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
    MS.Position := Entries.Items[i].LocalOffset;
    if not Entries.Items[i].ParseFromLocalHeader(MS) then
      FHasBadEntries := True;
  end;
end;

function TKAZip.ParseLocalHeaders(MS: TStream): Boolean;
var
  Poz: Integer;
  NLE: Integer;
  Sig: TZipSignature;
  Entry: TKAZipEntriesEntry;
  CDSize: Cardinal;
  NoMore: Boolean;
begin
  Result := False;
  FLocalHeaderNumFiles := 0;
  Entries.Clear();
  try
    Poz := 0;
    NLE := 0;
    CDSize := 0;
    repeat
      NoMore := True;
      MS.Position := Poz;
      MS.Read(Sig, Sizeof(Sig));
      if (Sig = SIG_LocalHeader) then
      begin
        Result := True;
        Inc(FLocalHeaderNumFiles);
        NoMore := False;
        MS.Position := Poz;

        Entry := TKAZipEntriesEntry.Create(Entries);
        if not Entry.ParseFromLocalHeader(MS) then
        begin
          Entry.Free();
        end;

        MS.Position := MS.Position + Entry.SizeCompressed;

        //FillCha r(CDFile, SizeOf(TCentralDirectoryFile), 0);  20140320 Cannot do FillCh ar on a structure containing managed types
        CDSize := CDSize + Entry.CentralEntrySize;
        Poz := MS.Position;
        Inc(NLE);
      end;
    until NoMore;

    FEndOfCentralDir.EndOfCentralDirSignature := SIG_EndOfCentralDirectory; //PK 0x05 0x06
    FEndOfCentralDir.NumberOfThisDisk := 0;
    FEndOfCentralDir.NumberOfTheDiskWithTheStart := 0;
    FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk := NLE;
    FEndOfCentralDir.SizeOfTheCentralDirectory := CDSize;
    FEndOfCentralDir.OffsetOfStartOfCentralDirectory := MS.Position;
    FEndOfCentralDir.ZipfileCommentLength := 0;
  except
    Exit;
  end;
end;

function TKAZip.AddStreamFast(ItemName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
var
  compressor: TZCompressionStream;
//  CS: TStringStream;
  compressionMode: Word;
//  S: string;
  i: Integer;
  uncompressedLength: Integer;
  compressedLength: Integer;
  dataCrc32: LongWord;
//  sizeToAppend: Integer;
  zipComment: string;
  compressionLevel: TZCompressionLevel;
  OBM: Boolean;
  crc32Stream: TCRC32Stream;
  newLocalEntryPosition: Int64;
  startOfCompressedDataPosition: Int64;
  centralDirectoryPosition: Int64;
  LFH: TLocalFileHeader;
begin
  //*********************************** COMPRESS DATA
  zipComment := Comment.Text;

  if not FStoreRelativePath then
    ItemName := ExtractFileName(ItemName);

  ItemName := ToZipName(ItemName);

  //If an item with this name already exists then remove it
  i := Entries.IndexOf(ItemName);
  if i >= 0 then
  begin
    OBM := FBatchMode;
    try
      if OBM = False then
        FBatchMode := True;
      Remove(i);
    finally
      FBatchMode := OBM;
    end;
  end;

  //This is where the new local entry starts (where the central directly is).
  //We overwrite the central direct (and EOCD marker) and then re-write them after the end of the file
  newLocalEntryPosition := FEndOfCentralDir.OffsetOfStartOfCentralDirectory;

  if Assigned(Stream) then
    uncompressedLength := Stream.Size - Stream.Position
  else
    uncompressedLength := 0;
  FCurrentDFS := uncompressedLength;


  dataCrc32 := 0; //we don't know it yet; back-fill it
//  compressedLength := 0; //we don't know it yet; back-fill it

  if uncompressedLength > 0 then
    compressionMode := ZipCompressionMethod_Deflate
  else
    compressionMode := ZipCompressionMethod_Store;

  //Fill records
  Result := TKAZipEntriesEntry(Entries.Add);
  Result.FileName := ItemName;
  Result.ExtraField := '';
  Result.Attributes := FileAttr;
  Result.Date := FileDate;
  Result.SizeUncompressed := uncompressedLength;
  Result.CompressionMethod := compressionMode;
  Result.LocalOffset := newLocalEntryPosition;

  //Local file entry
  Result.FillLocalHeader(LFH);

  // Write Local file entry to stream
  FZipStream.Position := newLocalEntryPosition;
  FZipStream.Write(LFH, SizeOf(LFH));
  if LFH.FilenameLength > 0 then
    FZipStream.Write(Result.FileName[1], LFH.FilenameLength);
  // extra field was set to empty

  //Now stream the compressed data
  startOfCompressedDataPosition := FZipStream.Position;

  if compressionMode = ZipCompressionMethod_Deflate then
  begin
    //The zlib compression level to use
    case FZipCompressionType of
      ctNormal: compressionLevel := zcDefault;
      ctMaximum: compressionLevel := zcMax;
      ctFast: compressionLevel := zcFastest;
      ctSuperFast: compressionLevel := zcFastest;
      ctNone: compressionLevel := zcNone;
      else
        compressionLevel := zcDefault;
    end;

    {
      We don't have access to a DEFLATE compressor.
      The best we have in ZLIB, which in turn uses DEFLATE.
      But zlib added 2 leading bytes, and 4 trailing bytes, that we need to strip off
        Skip first two zlib bytes  (RFC1950)
            CMF (Compression Method and flags)
            FLG (FLaGs)
        and final four bytes
            ADLER32 (Adler-32 checksum)
    }
//    OutputDebugString('SAMPLING ON');
    compressor := TZCompressionStream.Create(FZipStream, compressionLevel,
          -15, //windowBits - The default value is 15 if deflateInit is used instead.
          8, //memLevel - The default value is 8
          zsDefault); //strategy
    try
      compressor.OnProgress := OnCompress;

      //A through-stream that calculates the crc32 of the uncompressed data streaming through it
      crc32Stream := TCRC32Stream.Create(compressor);
      try
        crc32Stream.CopyFrom(Stream, uncompressedLength);
        dataCrc32 := crc32Stream.CRC32;
      finally
        crc32Stream.Free;
      end;
    finally
      compressor.Free;
    end;
//    OutputDebugString('SAMPLING OFF');
  end;

  //The central directory will start here; just after the file we just wrote
  //Memorize it now, because we're going to backtrack and update the CRC and compressed size
  //also because we can use it to figure out the compressed size
  centralDirectoryPosition := FZipStream.Position;

  if not Result.IsFolder then
  begin
    try
      compressedLength := (centralDirectoryPosition - startOfCompressedDataPosition);

      //Now that we know the CRC and the compressed size, update them in the local file entry
      Result.Crc32 := dataCrc32;
      Result.SizeCompressed := compressedLength;
      Result.FillLocalHeader(LFH);
      FZipStream.Position := newLocalEntryPosition;
      FZipStream.Write(LFH, SizeOf(LFH));
    finally
      //Return to where the central directory will be, and write it
      FZipStream.Position := centralDirectoryPosition;
    end;
  end;

  //Save the Central Directory entries and EOCD (End of Central Directory) record
  RebuildCentralDirectory(FZipStream);

  FIsDirty := True;
  if not FBatchMode then
  begin
    DoChange(Self, ZIP_ITEM_ADDED);
  end;
end;

function TKAZip.AddStreamRebuild(ItemName: string;
  FileAttr: Word;
  FileDate: TDateTime;
  Stream: TStream): TKAZipEntriesEntry;
var
  Compressor: TZCompressionStream;
  CS: TStringStream;
  cm: Word; //CompressionMethod to use
  S: AnsiString;
  UL: Integer;
  CL: Integer;
  I: Integer;
  X: Integer;
  FCRC32: Cardinal;
  OSL: Cardinal;
  NewSize: Cardinal;
  ZipComment: string;
  TempStream: TStream;
  TempFileName: string;
  Level: TZCompressionLevel;
  OBM: Boolean;
  LFH: TLocalFileHeader;
begin
  if FUseTempFiles then
  begin
    TempFileName := GetDelphiTempFileName;
    TempStream := TFileStream.Create(TempFileName, fmOpenReadWrite or FmCreate);
  end
  else
  begin
    TempStream := TMemoryStream.Create();
  end;

  try
    NewSize := 0;
    for X := 0 to Entries.Count - 1 do
    begin
      NewSize := NewSize + Entries.Items[X].LocalEntrySize + Entries.Items[X].CentralEntrySize;
      if Assigned(FOnRemoveItems) then
        FOnRemoveItems(Self, X, Entries.Count - 1);
    end;
    NewSize := NewSize + SizeOf(FEndOfCentralDir) + FEndOfCentralDir.ZipfileCommentLength;

    TempStream.Size := NewSize;
    TempStream.Position := 0;

    //*********************************** SAVE ALL OLD LOCAL ITEMS
    RebuildLocalFiles(TempStream);
    //*********************************** COMPRESS DATA
    ZipComment := Comment.Text;
    if not FStoreRelativePath then
      ItemName := ExtractFileName(ItemName);
    ItemName := ToZipName(ItemName);
    I := Entries.IndexOf(ItemName);
    if I > -1 then
    begin
      OBM := FBatchMode;
      try
        if OBM = False then
          FBatchMode := True;
        Remove(I);
      finally
        FBatchMode := OBM;
      end;
    end;

    cm := 0;
    CS := TStringStream.Create('');
    CS.Position := 0;
    try
      if Assigned(Stream) then
        UL := Stream.Size - Stream.Position
      else
        UL := 0;
      SetLength(S, UL);
      if UL > 0 then
      begin
        Stream.Read(S[1], UL);
        cm := ZipCompressionMethod_Deflate;
      end;
      FCRC32 := CalcCRC32(S);
      FCurrentDFS := UL;

      case FZipCompressionType of
        ctNormal: Level := zcDefault;
        ctMaximum: Level := zcMax;
        ctFast: Level := zcFastest;
        ctSuperFast: Level := zcFastest;
        ctNone: Level := zcNone;
        else
          Level := zcDefault;
      end;

      if cm = ZipCompressionMethod_Deflate then
      begin
        Compressor := TZCompressionStream.Create(CS, Level);
        try
          Compressor.OnProgress := OnCompress;
          Compressor.Write(S[1], UL);
        finally
          Compressor.Free;
        end;
        S := Copy(CS.DataString, 3, Length(CS.DataString) - 6);
      end;
    finally
      CS.Free;
    end;

    //************************************************************************
    CL := Length(S);
    //*********************************** FILL RECORDS
    Result := TKAZipEntriesEntry(Entries.Add);
    Result.Date := FileDate;
    Result.CompressionMethod := cm;
    Result.CRC32 := FCRC32;
    Result.SizeCompressed := CL;
    Result.SizeUncompressed := UL;
    Result.FileName := ItemName;
    Result.Attributes := FileAttr;
    Result.LocalOffset := TempStream.Position;

    //************************************ SAVE LOCAL HEADER AND COMPRESSED DATA
    Result.FillLocalHeader(LFH);
    TempStream.Write(LFH, SizeOf(LFH));
    if LFH.FilenameLength > 0 then
      TempStream.Write(Result.FileName[1], LFH.FilenameLength);
    if CL > 0 then
      TempStream.Write(S[1], CL);
    //************************************
    RebuildCentralDirectory(TempStream);
    //************************************
    TempStream.Position := 0;
    OSL := FZipStream.Size;
    try
      FZipStream.Size := TempStream.Size;
    except
      FZipStream.Size := OSL;
      raise;
    end;
    FZipStream.Position := 0;
    FZipStream.CopyFrom(TempStream, TempStream.Size);
  finally
    TempStream.Free();
    if FUseTempFiles then
      DeleteFile(TempFileName)
  end;

  FIsDirty := True;
  if not FBatchMode then
  begin
    DoChange(Self, ZIP_ITEM_ADDED);
  end;
end;

function TKAZip.AddFolderChain(ItemName: string; FileAttr: Word;
  FileDate: TDateTime): Boolean;
var
  FN: string;
  TN: string;
  INCN: string;
  P: Integer;
  NoMore: Boolean;
begin
  //  Result := False;
  FN := ExtractFilePath(ToDosName(ToZipName(ItemName)));
  TN := FN;
  INCN := '';
  repeat
    NoMore := True;
    P := Pos('\', TN);
    if P > 0 then
    begin
      INCN := INCN + Copy(TN, 1, P);
      System.Delete(TN, 1, P);
      if Entries.IndexOf(INCN) = -1 then
      begin
        if FZipSaveMethod = FastSave then
          AddStreamFast(INCN, FileAttr, FileDate, nil)
        else if FZipSaveMethod = RebuildAll then
          AddStreamRebuild(INCN, FileAttr, FileDate, nil);
      end;
      NoMore := False;
    end;
  until NoMore;
  Result := True;
end;

function TKAZip.AddFolderChain(ItemName: string): Boolean;
begin
  Result := AddFolderChain(ItemName, faDirectory, Now);
end;

function TKAZip.AddFolderEx(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
var
  Res: Integer;
  Dir: TSearchRec;
  FN: string;
begin
  Res := FindFirst(FolderName + '\*.*', faAnyFile, Dir);
  while Res = 0 do
  begin
    if (Dir.Attr and faDirectory) > 0 then
    begin
      if (Dir.Name <> '..') and (Dir.Name <> '.') then
      begin
        FN := FolderName + '\' + Dir.Name;
        if (FStoreFolders) and (FStoreRelativePath) then
          AddFolderChain(RemoveRootName(FN + '\', RootFolder), Dir.Attr, FileDateToDateTime(Dir.Time));
        if WithSubFolders then
        begin
          AddFolderEx(FN, RootFolder, WildCard, WithSubFolders);
        end;
      end
      else
      begin
        if (Dir.Name = '.') then
          AddFolderChain(RemoveRootName(FolderName + '\', RootFolder), Dir.Attr, FileDateToDateTime(Dir.Time));
      end;
    end
    else
    begin
      FN := FolderName + '\' + Dir.Name;
      if MatchesMask(FN, WildCard) then
      begin
        AddFile(FN, RemoveRootName(FN, RootFolder));
      end;
    end;
    Res := FindNext(Dir);
  end;
  FindClose(Dir);
  Result := True;
end;

function TKAZip.AddStream(FileName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := nil;
  if (FStoreFolders) and (FStoreRelativePath) then
    AddFolderChain(FileName);

  if FZipSaveMethod = FastSave then
    Result := AddStreamFast(FileName, FileAttr, FileDate, Stream)
  else if FZipSaveMethod = RebuildAll then
    Result := AddStreamRebuild(FileName, FileAttr, FileDate, Stream);

  if Assigned(FOnAddItem) then
    FOnAddItem(Self, FileName);
end;

function TKAZip.GetFileName(S: string): string;
var
  FN: string;
  P: Integer;
begin
  FN := S;
  FN := StringReplace(FN, '//', '\', [rfReplaceAll]);
  FN := StringReplace(FN, '/', '\', [rfReplaceAll]);
  P := Pos(':\', FN);
  if P > 0 then
    System.Delete(FN, 1, P + 1);
  Result := ExtractFileName(StringReplace(FN, '/', '\', [rfReplaceAll]));
end;

function TKAZip.AddStream(FileName: string; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := AddStream(FileName, faArchive, Now, Stream);
end;

function TKAZip.GetFilePath(S: string): string;
var
  FN: string;
  P: Integer;
begin
  FN := S;
  FN := StringReplace(FN, '//', '\', [rfReplaceAll]);
  FN := StringReplace(FN, '/', '\', [rfReplaceAll]);
  P := Pos(':\', FN);
  if P > 0 then
    System.Delete(FN, 1, P + 1);
  Result := ExtractFilePath(StringReplace(FN, '/', '\', [rfReplaceAll]));
end;

procedure TKAZip.LoadFromFile(FileName: string);
var
  Res: Integer;
  Dir: TSearchRec;
begin
  Res := FindFirst(FileName, faAnyFile, Dir);
  if Res = 0 then
  begin
    if Dir.Attr and faReadOnly > 0 then
    begin
      FZipStream := TFileStream.Create(FileName, fmOpenRead or fmShareDenyNone);
      FReadOnly := True;
    end
    else
    begin
      FZipStream := TFileStream.Create(FileName, fmOpenReadWrite or fmShareDenyNone);
      FReadOnly := False;
    end;
    LoadFromStream(FZipStream);
  end
  else
  begin
    raise Exception.Create('File "' + FileName + '" not found!');
  end;
end;

procedure TKAZip.LoadFromStream(MS: TStream);
begin
  FZipStream := MS;
  ParseZip(MS);
  if not FIsZipFile then
    Close;
  FIsDirty := True;
  DoChange(Self, ZIP_OPENED);
end;

procedure TKAZip.Close;
begin
  Entries.Clear;
  if Assigned(FZipStream) and (not FExternalStream) then
    FZipStream.Free;
  FExternalStream := False;
  FZipStream := nil;
  FIsZipFile := False;
  FIsDirty := True;
  FReadOnly := False;
  DoChange(Self, ZIP_CLOSED);
end;

procedure TKAZip.SetFileName(const Value: string);
begin
  FFileName := Value;
end;

procedure TKAZip.Open(FileName: string);
begin
  Close;
  LoadFromFile(FileName);
  FFileName := FileName;
end;

procedure TKAZip.Open(MS: TStream);
begin
  try
    Close;
    LoadFromStream(MS);
  finally
    FExternalStream := True;
  end;
end;

procedure TKAZip.SetIsZipFile(const Value: Boolean);
begin
  //****************************************************************************
end;

function TKAZip.GetDelphiTempFileName: string;
var
  TmpDir: array[0..1000] of Char;
  TmpFN: array[0..1000] of Char;
begin
  Result := GetCurrentDir;
  if GetTempPath(1000, TmpDir) <> 0 then
  begin
    if GetTempFileName(TmpDir, '', 0, TmpFN) <> 0 then
      Result := StrPas(TmpFN);
  end;
end;

procedure TKAZip.OnDecompress(Sender: TObject);
var
  DS: TStream;
begin
  DS := TStream(Sender);
  if Assigned(FOnDecompressFile) then
    FOnDecompressFile(Self, DS.Position, FCurrentDFS);
end;

procedure TKAZip.OnCompress(Sender: TObject);
var
  CS: TStream;
begin
  CS := TStream(Sender);
  if Assigned(FOnCompressFile) then
    FOnCompressFile(Self, CS.Position, FCurrentDFS);
end;

procedure TKAZip.ExtractToFile(Item: TKAZipEntriesEntry; FileName: string);
var
  Can: Boolean;
  OA: TOverwriteAction;
begin
  OA := FOverwriteAction;
  Can := True;
  if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FOnOverwriteFile)) then
  begin
    if FileExists(FileName) then
    begin
      FOnOverwriteFile(Self, FileName, OA);
    end
    else
    begin
      OA := oaOverwrite;
    end;
  end;
  case OA of
    oaSkip: Can := False;
    oaSkipAll: Can := False;
    oaOverwrite: Can := True;
    oaOverwriteAll: Can := True;
  end;
  if Can then
    InternalExtractToFile(Item, FileName);
end;

procedure TKAZip.ExtractToFile(ItemIndex: Integer; FileName: string);
var
  Can: Boolean;
  OA: TOverwriteAction;
begin
  OA := FOverwriteAction;
  Can := True;
  if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FOnOverwriteFile)) then
  begin
    if FileExists(FileName) then
    begin
      FOnOverwriteFile(Self, FileName, OA);
    end
    else
    begin
      OA := oaOverwrite;
    end;
  end;
  case OA of
    oaSkip: Can := False;
    oaSkipAll: Can := False;
    oaOverwrite: Can := True;
    oaOverwriteAll: Can := True;
  end;
  if Can then
    InternalExtractToFile(Entries.Items[ItemIndex], FileName);
end;

procedure TKAZip.ExtractToFile(FileName, DestinationFileName: string);
var
  I: Integer;
  Can: Boolean;
  OA: TOverwriteAction;
begin
  OA := FOverwriteAction;
  Can := True;
  if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FOnOverwriteFile)) then
  begin
    if FileExists(DestinationFileName) then
    begin
      FOnOverwriteFile(Self, DestinationFileName, OA);
    end
    else
    begin
      OA := oaOverwrite;
    end;
  end;
  case OA of
    oaSkip: Can := False;
    oaSkipAll: Can := False;
    oaOverwrite: Can := True;
    oaOverwriteAll: Can := True;
  end;
  if Can then
  begin
    I := Entries.IndexOf(FileName);
    InternalExtractToFile(Entries.Items[I], DestinationFileName);
  end;
end;

procedure TKAZip.ExtractToStream(Item: TKAZipEntriesEntry; TargetStream: TStream);
var
  sourceStream: TMemoryStream;
  buf: AnsiString;
  NR: Cardinal;
  decompressor: TZDecompressionStream;
{$IFDEF USE_BZIP2}
  DecompressorBZ2: TBZDecompressionStream;
{$ENDIF}
begin
  if (Item.FIsEncrypted) then
    raise Exception.Create('Cannot process file "' + Item.FileName + '": File is encrypted');

  if (Item.CompressionMethod <> ZipCompressionMethod_Deflate)
{$IFDEF USE_BZIP2}
    and (Item.CompressionMethod <> ZipCompressionMethod_bzip2)
{$ENDIF}
    and (Item.CompressionMethod <> ZipCompressionMethod_Store) then
  begin
    raise Exception.Create('Cannot process file "' + Item.FileName + '": Unknown compression method '+IntToStr(Item.CompressionMethod));
  end;

  sourceStream := TMemoryStream.Create;
  try
    if Item.GetCompressedData(sourceStream) <= 0 then
      Exit;

    sourceStream.Position := 0;
    FCurrentDFS := Item.SizeUncompressed;

    case Item.CompressionMethod of
      ZipCompressionMethod_Deflate:
        begin
          decompressor := TZDecompressionStream.Create(sourceStream);
          decompressor.OnProgress := OnDecompress;
          SetLength(BUF, FCurrentDFS);
          try
            NR := Decompressor.Read(BUF[1], FCurrentDFS);
            if NR = FCurrentDFS then
              TargetStream.Write(BUF[1], FCurrentDFS);
          finally
            decompressor.Free;
          end;
        end;
{$IFDEF USE_BZIP2}
      ZipCompressionMethod_bzip2:
        begin
          DecompressorBZ2 := TBZDecompressionStream.Create(sourceStream);
          DecompressorBZ2.OnProgress := OnDecompress;
          SetLength(BUF, FCurrentDFS);
          try
            NR := DecompressorBZ2.Read(BUF[1], FCurrentDFS);
            if NR = FCurrentDFS then
              TargetStream.Write(BUF[1], FCurrentDFS);
          finally
            DecompressorBZ2.Free;
          end;
        end
{$ENDIF}
      ZipCompressionMethod_Store:
        begin
          TargetStream.CopyFrom(sourceStream, FCurrentDFS);
        end;
    end;
  finally
    sourceStream.Free;
  end;
end;

procedure TKAZip.ExtractAll(TargetDirectory: string);
var
  FN: string;
  DN: string;
  X: Integer;
  Can: Boolean;
  OA: TOverwriteAction;
  FileName: string;
begin
  OA := FOverwriteAction;
  Can := True;
  try
    for X := 0 to Entries.Count - 1 do
    begin
      FN := GetFileName(Entries.Items[X].FileName);
      DN := GetFilePath(Entries.Items[X].FileName);
      if DN <> '' then
        ForceDirectories(TargetDirectory + '\' + DN);
      FileName := TargetDirectory + '\' + DN + FN;
      if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FOnOverwriteFile)) then
      begin
        if FileExists(FileName) then
        begin
          FOnOverwriteFile(Self, FileName, OA);
        end;
      end;
      case OA of
        oaSkip: Can := False;
        oaSkipAll: Can := False;
        oaOverwrite: Can := True;
        oaOverwriteAll: Can := True;
      end;
      if Can then
        InternalExtractToFile(Entries.Items[X], FileName);
    end;
  finally
  end;
end;

procedure TKAZip.ExtractSelected(TargetDirectory: string);
var
  FN: string;
  DN: string;
  X: Integer;
  OA: TOverwriteAction;
  Can: Boolean;
  FileName: string;
begin
  OA := FOverwriteAction;
  Can := True;
  try
    for X := 0 to Entries.Count - 1 do
    begin
      if Entries.Items[X].Selected then
      begin
        FN := GetFileName(Entries.Items[X].FileName);
        DN := GetFilePath(Entries.Items[X].FileName);
        if DN <> '' then
          ForceDirectories(TargetDirectory + '\' + DN);
        FileName := TargetDirectory + '\' + DN + FN;
        if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FOnOverwriteFile)) then
        begin
          if FileExists(FileName) then
          begin
            FOnOverwriteFile(Self, FileName, OA);
          end;
        end;
        case OA of
          oaSkip: Can := False;
          oaSkipAll: Can := False;
          oaOverwrite: Can := True;
          oaOverwriteAll: Can := True;
        end;
        if Can then
          InternalExtractToFile(Entries.Items[X], TargetDirectory + '\' + DN + FN);
      end;
    end;
  finally
  end;
end;

function TKAZip.AddFile(FileName, NewFileName: string): TKAZipEntriesEntry;
var
  FS: TFileStream;
  Dir: TSearchRec;
  Res: Integer;
begin
  Result := nil;
  Res := FindFirst(FileName, faAnyFile, Dir);
  if Res = 0 then
  begin
    FS := TFileStream.Create(FileName, fmOpenRead or fmShareDenyNone);
    try
      FS.Position := 0;
      Result := AddStream(NewFileName, Dir.Attr, FileDateToDateTime(Dir.Time), FS)
    finally
      FS.Free;
    end;
  end;
  FindClose(Dir);
end;

function TKAZip.AddFile(FileName: string): TKAZipEntriesEntry;
begin
  Result := Entries.AddFile(FileName);
end;

function TKAZip.AddFiles(FileNames: TStrings): Boolean;
var
  X: Integer;
begin
  Result := False;
  FBatchMode := True;
  try
    for X := 0 to FileNames.Count - 1 do
      AddFile(FileNames.Strings[X]);
  except
    FBatchMode := False;
    DoChange(Self, ZIP_ITEM_ADDED);
    Exit;
  end;
  FBatchMode := False;
  DoChange(Self, ZIP_ITEM_ADDED);
  Result := True;
end;

function TKAZip.AddFolder(FolderName, RootFolder, WildCard: string;
  WithSubFolders: Boolean): Boolean;
begin
  FBatchMode := True;
  try
    Result := AddFolderEx(FolderName, RootFolder, WildCard, WithSubFolders);
  finally
    FBatchMode := False;
    DoChange(Self, ZIP_ITEM_ADDED);
  end;
end;

function TKAZip.AddFilesAndFolders(FileNames: TStrings; RootFolder: string;
  WithSubFolders: Boolean): Boolean;
var
  X: Integer;
  Res: Integer;
  Dir: TSearchRec;
begin
  FBatchMode := True;
  try
    for X := 0 to FileNames.Count - 1 do
    begin
      Res := FindFirst(FileNames.Strings[X], faAnyFile, Dir);
      if Res = 0 then
      begin
        if (Dir.Attr and faDirectory) > 0 then
        begin
          if (Dir.Name <> '..') and (Dir.Name <> '.') then
          begin
            AddFolderEx(FileNames.Strings[X], RootFolder, '*.*', WithSubFolders);
          end;
        end
        else
        begin
          AddFile(FileNames.Strings[X], RemoveRootName(FileNames.Strings[X], RootFolder));
        end;
      end;
      FindClose(Dir);
    end;
  finally
    FBatchMode := False;
    DoChange(Self, ZIP_ITEM_ADDED);
  end;
  Result := True;
end;

procedure TKAZip.Remove(ItemIndex: Integer; Flush: Boolean);
var
  TempStream: TFileStream;
  TempMSStream: TMemoryStream;
  TempFileName: string;
//  BUF: string;
  ZipComment: string;
  OSL: Cardinal;
  //*********************************************
  X: Integer;
  TargetPos: Cardinal;
  Border: Cardinal;

//  NR: Integer;
//  NW: Integer;
//  BufStart: Integer;
//  BufLen: Integer;
  ShiftSize: Cardinal;
  NewSize: Cardinal;
  CDFile: TCentralDirectoryFile;
begin
  TargetPos := Entries.Items[ItemIndex].LocalOffset;
  ShiftSize := Entries.Items[ItemIndex].LocalEntrySize;
//  BufStart := TargetPos + ShiftSize;
//  BufLen := FZipStream.Size - BufStart;
  Border := TargetPos;
  Entries.Delete(ItemIndex);
  if (FZipSaveMethod = FastSave) and (Entries.Count > 0) then
  begin
    ZipComment := Comment.Text;

{
    SetLength(BUF,BufLen);
    FZipStream.Position := BufStart;
    NR := FZipStream.Read(BUF[1],BufLen);

    FZipStream.Position := TargetPos;
    NW := FZipStream.Write(BUF[1],BufLen);
    SetLength(BUF,0);
  }

    for X := 0 to Entries.Count - 1 do
    begin
      if Entries.Items[X].LocalOffset > Border then
      begin
        Entries.Items[X].LocalOffset := Entries.Items[X].LocalOffset - ShiftSize;
        TargetPos := TargetPos + Entries.Items[X].LocalEntrySize;
      end
    end;

    FZipStream.Position := TargetPos;
    //************************************ MARK START OF CENTRAL DIRECTORY
    FEndOfCentralDir.OffsetOfStartOfCentralDirectory := FZipStream.Position;
    //************************************ SAVE CENTRAL DIRECTORY
    for X := 0 to Entries.Count - 1 do
    begin
      Entries.Items[X].FillCentralDirectory(CDFile);
      FZipStream.Write(CDFile, SizeOf(CDFile));
      if CDFile.FilenameLength > 0 then
        FZipStream.Write(Entries.Items[X].FileName[1], CDFile.FilenameLength);
      if CDFile.ExtraFieldLength > 0 then
        FZipStream.Write(Entries.Items[X].ExtraField[1], CDFile.ExtraFieldLength);
      if CDFile.FileCommentLength > 0 then
        FZipStream.Write(Entries.Items[X].Comment[1], CDFile.FileCommentLength);
    end;
    //************************************ SAVE END CENTRAL DIRECTORY RECORD
    FEndOfCentralDirPos := FZipStream.Position;
    FEndOfCentralDir.SizeOfTheCentralDirectory := FEndOfCentralDirPos - FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
    Dec(FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
    Dec(FEndOfCentralDir.TotalNumberOfEntries);
    FZipStream.Write(FEndOfCentralDir, SizeOf(TEndOfCentralDir));
    //************************************ SAVE ZIP COMMENT IF ANY
    FZipCommentPos := FZipStream.Position;
    if Length(ZipComment) > 0 then
    begin
      FZipStream.Write(ZipComment[1], Length(ZipComment));
    end;
    FZipStream.Size := FZipStream.Position;
  end
  else
  begin
    if FUseTempFiles then
    begin
      TempFileName := GetDelphiTempFileName;
      TempStream := TFileStream.Create(TempFileName, fmOpenReadWrite or FmCreate);
      try
        SaveToStream(TempStream);
        TempStream.Position := 0;
        OSL := FZipStream.Size;
        try
          FZipStream.Size := TempStream.Size;
        except
          FZipStream.Size := OSL;
          raise;
        end;
        FZipStream.Position := 0;
        FZipStream.CopyFrom(TempStream, TempStream.Size);
        //*********************************************************************
        FZipHeader.ParseZip(FZipStream);
        //*********************************************************************
      finally
        TempStream.Free;
        DeleteFile(TempFileName)
      end;
    end
    else
    begin
      NewSize := 0;
      for X := 0 to Entries.Count - 1 do
      begin
        NewSize := NewSize + Entries.Items[X].LocalEntrySize + Entries.Items[X].CentralEntrySize;
        if Assigned(FOnRemoveItems) then
          FOnRemoveItems(Self, X, Entries.Count - 1);
      end;
      NewSize := NewSize + SizeOf(FEndOfCentralDir) + FEndOfCentralDir.ZipfileCommentLength;
      TempMSStream := TMemoryStream.Create;
      try
        TempMSStream.SetSize(NewSize);
        TempMSStream.Position := 0;
        SaveToStream(TempMSStream);
        TempMSStream.Position := 0;
        OSL := FZipStream.Size;
        try
          FZipStream.Size := TempMSStream.Size;
        except
          FZipStream.Size := OSL;
          raise;
        end;
        FZipStream.Position := 0;
        FZipStream.CopyFrom(TempMSStream, TempMSStream.Size);
        //*********************************************************************
        FZipHeader.ParseZip(FZipStream);
        //*********************************************************************
      finally
        TempMSStream.Free;
      end;
    end;
  end;
  FIsDirty := True;
  if not FBatchMode then
  begin
    DoChange(Self, ZIP_ITEM_REMOVED);
  end;
end;

procedure TKAZip.RemoveBatch(Files: TList);
var
  X: Integer;
  OSL: Integer;
  NewSize: Cardinal;
  TempStream: TFileStream;
  TempMSStream: TMemoryStream;
  TempFileName: string;
begin
  for X := Files.Count - 1 downto 0 do
  begin
    Entries.Delete(Integer(Files.Items[X]));
    if Assigned(FOnRemoveItems) then
      FOnRemoveItems(Self, Files.Count - X, Files.Count);
  end;
  NewSize := 0;
  if FUseTempFiles then
  begin
    TempFileName := GetDelphiTempFileName;
    TempStream := TFileStream.Create(TempFileName, fmOpenReadWrite or FmCreate);
    try
      SaveToStream(TempStream);
      TempStream.Position := 0;
      OSL := FZipStream.Size;
      try
        FZipStream.Size := TempStream.Size;
      except
        FZipStream.Size := OSL;
        raise;
      end;
      FZipStream.Position := 0;
      FZipStream.CopyFrom(TempStream, TempStream.Size);
      //*********************************************************************
      FZipHeader.ParseZip(FZipStream);
      //*********************************************************************
    finally
      TempStream.Free;
      DeleteFile(TempFileName)
    end;
  end
  else
  begin
    for X := 0 to Entries.Count - 1 do
    begin
      NewSize := NewSize + Entries.Items[X].LocalEntrySize + Entries.Items[X].CentralEntrySize;
      if Assigned(FOnRemoveItems) then
        FOnRemoveItems(Self, X, Entries.Count - 1);
    end;
    NewSize := NewSize + SizeOf(FEndOfCentralDir) + FEndOfCentralDir.ZipfileCommentLength;
    TempMSStream := TMemoryStream.Create;
    try
      TempMSStream.SetSize(NewSize);
      TempMSStream.Position := 0;
      SaveToStream(TempMSStream);
      TempMSStream.Position := 0;
      OSL := FZipStream.Size;
      try
        FZipStream.Size := TempMSStream.Size;
      except
        FZipStream.Size := OSL;
        raise;
      end;
      FZipStream.Position := 0;
      FZipStream.CopyFrom(TempMSStream, TempMSStream.Size);
      //*********************************************************************
      FZipHeader.ParseZip(FZipStream);
      //*********************************************************************
    finally
      TempMSStream.Free;
    end;
  end;
end;

procedure TKAZip.InternalExtractToFile(Item: TKAZipEntriesEntry; FileName: string);
var
  TFS: TFileStream;
  Attr: Integer;
begin
  if Item.IsFolder then
  begin
    ForceDirectories(FileName);
  end
  else
  begin
    TFS := TFileStream.Create(FileName, fmCreate or fmOpenReadWrite or fmShareDenyNone);
    try
      ExtractToStream(Item, TFS);
    finally
      TFS.Free;
    end;
    if FApplyAttributes then
    begin
      Attr := faArchive;
      if Item.Attributes and faHidden > 0 then
        Attr := Attr or faHidden;
      if Item.Attributes and faSysFile > 0 then
        Attr := Attr or faSysFile;
      if Item.Attributes and faReadOnly > 0 then
        Attr := Attr or faReadOnly;
      FileSetAttr(FileName, Attr);
    end;
  end;
end;

procedure TKAZip.Remove(Item: TKAZipEntriesEntry);
begin
  Entries.Remove(Item);
end;

procedure TKAZip.Remove(ItemIndex: Integer);
begin
  Entries.Remove(ItemIndex);
end;

procedure TKAZip.Remove(FileName: string);
begin
  Entries.Remove(FileName);
end;

procedure TKAZip.RemoveFiles(List: TList);
begin
  if List.Count = 1 then
  begin
    Remove(Integer(List.Items[0]));
  end
  else
  begin
    Entries.SortList(List);
    FBatchMode := True;
    try
      RemoveBatch(List);
    finally
      FBatchMode := False;
      DoChange(Self, ZIP_ITEM_REMOVED);
    end;
  end;
end;

procedure TKAZip.RemoveSelected;
var
  X: Integer;
  List: TList;
begin
  FBatchMode := True;
  List := TList.Create;
  try
    for X := 0 to Entries.Count - 1 do
    begin
      if Entries.Items[X].Selected then
        List.Add(Pointer(X));
    end;
    RemoveBatch(List);
  finally
    List.Free;
    FBatchMode := False;
    DoChange(Self, ZIP_ITEM_REMOVED);
  end;
end;

function TKAZip.GetComment: TStrings;
var
  S: string;
begin
  Result := FZipComment;
  FZipComment.Clear; //20131104 - Why would you clear the zip comment just because i read it?
  if FIsZipFile then
  begin
    if FEndOfCentralDir.ZipfileCommentLength > 0 then
    begin
      FZipStream.Position := FZipCommentPos;
      SetLength(S, FEndOfCentralDir.ZipfileCommentLength);
      FZipStream.Read(S[1], FEndOfCentralDir.ZipfileCommentLength);
      FZipComment.Text := S;
    end;
  end;
end;

procedure TKAZip.SetComment(const Value: TStrings);
var
  Comment: string;
  L: Integer;
begin
  //****************************************************************************
  if FZipComment.Text = Value.Text then
    Exit;
  FZipComment.Clear;
  if FIsZipFile then
  begin
    FZipComment.Assign(Value);
    Comment := FZipComment.Text;
    L := Length(Comment);
    FEndOfCentralDir.ZipfileCommentLength := L;
    FZipStream.Position := FEndOfCentralDirPos;
    FZipStream.Write(FEndOfCentralDir, SizeOf(TEndOfCentralDir));
    FZipCommentPos := FZipStream.Position;
    if L > 0 then
    begin
      FZipStream.Write(Comment[1], L)
    end
    else
    begin
      FZipStream.Size := FZipStream.Position;
    end;
  end;
end;

procedure TKAZip.DeSelectAll;
begin
  Entries.DeSelectAll;
end;

procedure TKAZip.Select(WildCard: string);
begin
  Entries.Select(WildCard);
end;

procedure TKAZip.InvertSelection;
begin
  Entries.InvertSelection;
end;

procedure TKAZip.SelectAll;
begin
  Entries.SelectAll;
end;

procedure TKAZip.RebuildLocalFiles(MS: TStream);
var
  X: Integer;
  Entry: TKAZipEntriesEntry;
  LFH: TLocalFileHeader;
begin
  //************************************************* RESAVE ALL LOCAL BLOCKS
  for X := 0 to Entries.Count - 1 do
  begin
    Entry := Entries.Items[X];
    Entry.LocalOffset := MS.Position;
    Entry.FillLocalHeader(LFH);
    MS.Write(LFH, SizeOf(LFH));
    if LFH.FilenameLength > 0 then
      MS.Write(Entry.FileName[1], LFH.FilenameLength);
    if LFH.ExtraFieldLength > 0 then
      MS.Write(Entry.ExtraField[1], LFH.ExtraFieldLength);
    if LFH.CompressedSize > 0 then
    begin
      FZipStream.Position := Entry.CompressedDataPos;
      MS.CopyFrom(FZipStream, LFH.CompressedSize);
    end;
    if Assigned(FOnRebuildZip) then
      FOnRebuildZip(Self, X, Entries.Count - 1);
  end;
end;

procedure TKAZip.RebuildCentralDirectory(MS: TStream);
var
  i: Integer;
  Entry: TKAZipEntriesEntry;
  CDF: TCentralDirectoryFile;
  ZipComment: string;
begin
  FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk := Entries.Count;
  FEndOfCentralDir.TotalNumberOfEntries := Entries.Count;
  FEndOfCentralDir.OffsetOfStartOfCentralDirectory := MS.Position;
  for i := 0 to Entries.Count - 1 do
  begin
    Entry := Entries.Items[i];
    Entry.FillCentralDirectory(CDF);
    MS.Write(CDF, SizeOf(CDF));
    if CDF.FilenameLength > 0 then
      MS.Write(Entry.FileName[1], CDF.FilenameLength);
    if CDF.ExtraFieldLength > 0 then
      MS.Write(Entry.ExtraField[1], CDF.ExtraFieldLength);
    if CDF.FileCommentLength > 0 then
      MS.Write(Entry.Comment[1], CDF.FileCommentLength);
    if Assigned(FOnRebuildZip) then
      FOnRebuildZip(Self, i, Entries.Count - 1);
  end;
  FEndOfCentralDir.SizeOfTheCentralDirectory := Cardinal(MS.Position) - FEndOfCentralDir.OffsetOfStartOfCentralDirectory;

  // Rebuild end of Central Directory
  ZipComment := Comment.Text;
  FEndOfCentralDir.ZipfileCommentLength := Length(ZipComment);
  FEndOfCentralDirPos := FZipStream.Position;
  FRebuildECDP := MS.Position;
  MS.Write(FEndOfCentralDir, SizeOf(FEndOfCentralDir));
  FRebuildCP := MS.Position;
  if FEndOfCentralDir.ZipfileCommentLength > 0 then
  begin
    FZipCommentPos := FZipStream.Position;
    MS.Write(ZipComment[1], FEndOfCentralDir.ZipfileCommentLength);
  end;
  if Assigned(FOnRebuildZip) then
    FOnRebuildZip(Self, 100, 100);
end;

procedure TKAZip.FixZip(MS: TStream);
var
  X: Integer;
  Y: Integer;
  NewCount: Integer;
  Entry: TKAZipEntriesEntry;
  LFH: TLocalFileHeader;
  CDF: TCentralDirectoryFile;
  ZipComment: string;
  NewLHOffsets: array of Cardinal;
  NewEndOfCentralDir: TEndOfCentralDir;
begin
  ZipComment := Comment.Text;
  Y := 0;
  SetLength(NewLHOffsets, Entries.Count + 1);
  for X := 0 to Entries.Count - 1 do
  begin
    Entry := Entries.Items[X];
    if Entry.Test() then
    begin
      Entry.FillLocalHeader(LFH);
      // good entry positions in new zip
      NewLHOffsets[Y] := MS.Position;
      MS.Write(LFH, SizeOf(LFH));
      if LFH.FilenameLength > 0 then
        MS.Write(Entry.FileName[1], LFH.FilenameLength);
      if LFH.ExtraFieldLength > 0 then
        MS.Write(Entry.ExtraField[1], LFH.ExtraFieldLength);
      if LFH.CompressedSize > 0 then
      begin
        FZipStream.Position := Entry.CompressedDataPos;
        MS.CopyFrom(FZipStream, LFH.CompressedSize);
      end;
      if Assigned(FOnRebuildZip) then
        FOnRebuildZip(Self, X, Entries.Count - 1);
      Inc(Y);
    end
    else
    begin
      if Assigned(FOnRebuildZip) then
        FOnRebuildZip(Self, X, Entries.Count - 1);
    end;
  end;

  NewCount := Y;
  Y := 0;
  NewEndOfCentralDir := FEndOfCentralDir;
  NewEndOfCentralDir.TotalNumberOfEntriesOnThisDisk := NewCount;
  NewEndOfCentralDir.TotalNumberOfEntries := NewCount;
  NewEndOfCentralDir.OffsetOfStartOfCentralDirectory := MS.Position;
  for X := 0 to Entries.Count - 1 do
  begin
    Entry := Entries.Items[X];
    Entry.FillCentralDirectory(CDF);
    // set entry position for new zip
    CDF.RelativeOffsetOfLocalHeader := NewLHOffsets[Y];
    MS.Write(CDF, SizeOf(CDF));
    if CDF.FilenameLength > 0 then
      MS.Write(Entry.FileName[1], CDF.FilenameLength);
    if CDF.ExtraFieldLength > 0 then
      MS.Write(Entry.ExtraField[1], CDF.ExtraFieldLength);
    if CDF.FileCommentLength > 0 then
      MS.Write(Entry.Comment[1], CDF.FileCommentLength);
    if Assigned(FOnRebuildZip) then
      FOnRebuildZip(Self, X, Entries.Count - 1);
    Inc(Y);
  end;
  NewEndOfCentralDir.SizeOfTheCentralDirectory := Cardinal(MS.Position) - NewEndOfCentralDir.OffsetOfStartOfCentralDirectory;

  FRebuildECDP := MS.Position;
  MS.Write(NewEndOfCentralDir, SizeOf(NewEndOfCentralDir));
  FRebuildCP := MS.Position;
  if NewEndOfCentralDir.ZipfileCommentLength > 0 then
  begin
    MS.Write(ZipComment[1], NewEndOfCentralDir.ZipfileCommentLength);
  end;
  if Assigned(FOnRebuildZip) then
    FOnRebuildZip(Self, 100, 100);
end;

procedure TKAZip.SaveToStream(Stream: TStream);
begin
  RebuildLocalFiles(Stream);
  RebuildCentralDirectory(Stream);
end;

procedure TKAZip.Rebuild;
var
  TempStream: TFileStream;
  TempMSStream: TMemoryStream;
  TempFileName: string;
begin
  if FUseTempFiles then
  begin
    TempFileName := GetDelphiTempFileName;
    TempStream := TFileStream.Create(TempFileName, fmOpenReadWrite or FmCreate);
    try
      SaveToStream(TempStream);
      FZipStream.Position := 0;
      FZipStream.Size := 0;
      TempStream.Position := 0;
      FZipStream.CopyFrom(TempStream, TempStream.Size);
      Entries.ParseZip(FZipStream);
    finally
      TempStream.Free;
      DeleteFile(TempFileName)
    end;
  end
  else
  begin
    TempMSStream := TMemoryStream.Create;
    try
      SaveToStream(TempMSStream);
      FZipStream.Position := 0;
      FZipStream.Size := 0;
      TempMSStream.Position := 0;
      FZipStream.CopyFrom(TempMSStream, TempMSStream.Size);
      Entries.ParseZip(FZipStream);
    finally
      TempMSStream.Free;
    end;
  end;
  FIsDirty := True;
end;

procedure TKAZip.CreateZip(Stream: TStream);
var
  ECD: TEndOfCentralDir;
begin
  ECD.EndOfCentralDirSignature := SIG_EndOfCentralDirectory;
  ECD.NumberOfThisDisk := 0;
  ECD.NumberOfTheDiskWithTheStart := 0;
  ECD.TotalNumberOfEntriesOnThisDisk := 0;
  ECD.TotalNumberOfEntries := 0;
  ECD.SizeOfTheCentralDirectory := 0;
  ECD.OffsetOfStartOfCentralDirectory := 0;
  ECD.ZipfileCommentLength := 0;
  Stream.Write(ECD, SizeOf(ECD));
end;

procedure TKAZip.CreateZip(FileName: string);
var
  FS: TFileStream;
begin
  FS := TFileStream.Create(FileName, fmOpenReadWrite or FmCreate);
  try
    CreateZip(FS);
  finally
    FS.Free;
  end;
end;

procedure TKAZip.SetZipSaveMethod(const Value: TZipSaveMethod);
begin
  FZipSaveMethod := Value;
end;

procedure TKAZip.SetActive(const Value: Boolean);
begin
  if FFileName = '' then
    Exit;
  if Value then
    Open(FFileName)
  else
    Close;
end;

procedure TKAZip.SetZipCompressionType(const Value: TZipCompressionType);
begin
  FZipCompressionType := Value;
  if FZipCompressionType = ctUnknown then
    FZipCompressionType := ctNormal;
end;

function TKAZip.GetFileNames: TStrings;
var
  X: Integer;
begin
  if FIsDirty then
  begin
    FFileNames.Clear;
    for X := 0 to Entries.Count - 1 do
    begin
      FFileNames.Add(GetFilePath(Entries.Items[X].FileName) + GetFileName(Entries.Items[X].FileName));
    end;
    FIsDirty := False;
  end;
  Result := FFileNames;
end;

procedure TKAZip.SetFileNames(const Value: TStrings);
begin
  //*************************************************** READ ONLY
end;

procedure TKAZip.SetUseTempFiles(const Value: Boolean);
begin
  FUseTempFiles := Value;
end;

procedure TKAZip.Rename(Item: TKAZipEntriesEntry; NewFileName: string);
begin
  Entries.Rename(Item, NewFileName);
end;

procedure TKAZip.Rename(ItemIndex: Integer; NewFileName: string);
begin
  Entries.Rename(ItemIndex, NewFileName);
end;

procedure TKAZip.Rename(FileName, NewFileName: string);
begin
  Entries.Rename(FileName, NewFileName);
end;

procedure TKAZip.RenameMultiple(Names, NewNames: TStringList);
var
  X: Integer;
  BR: Integer;
  L: Integer;
begin
  BR := 0;

  if Names.Count <> NewNames.Count then
  begin
    raise Exception.Create('Names and NewNames must have equal count');
  end
  else
  begin
    FBatchMode := True;
    try
      for X := 0 to Names.Count - 1 do
      begin
        L := Length(Names.Strings[X]);
        if (L > 0) and ((Names.Strings[X][L] = '\') or (Names.Strings[X][L] = '/')) then
        begin
          RenameFolder(Names.Strings[X], NewNames.Strings[X]);
          Inc(BR);
        end
        else
        begin
          Rename(Names.Strings[X], NewNames.Strings[X]);
          Inc(BR);
        end;
      end;
    finally
      FBatchMode := False;
    end;
    if BR > 0 then
    begin
      Rebuild();
      DoChange(Self, ZIP_ITEMS_NAME);
    end;
  end;
end;

procedure TKAZip.SetStoreFolders(const Value: Boolean);
begin
  FStoreFolders := Value;
end;

procedure TKAZip.SetOnAddItem(const Value: TOnAddItem);
begin
  FOnAddItem := Value;
end;

procedure TKAZip.SetComponentVersion(const Value: string);
begin
  //****************************************************************************
end;

procedure TKAZip.SetOnRebuildZip(const Value: TOnRebuildZip);
begin
  FOnRebuildZip := Value;
end;

procedure TKAZip.SetOnRemoveItems(const Value: TOnRemoveItems);
begin
  FOnRemoveItems := Value;
end;

procedure TKAZip.SetOverwriteAction(const Value: TOverwriteAction);
begin
  FOverwriteAction := Value;
end;

procedure TKAZip.SetOnOverwriteFile(const Value: TOnOverwriteFile);
begin
  FOnOverwriteFile := Value;
end;

procedure TKAZip.CreateFolder(FolderName: string; FolderDate: TDateTime);
var
  FN: string;
begin
  FN := IncludeTrailingPathDelimiter(FolderName);
  AddFolderChain(FN, faDirectory, FolderDate);
  FIsDirty := True;
end;

procedure TKAZip.RenameFolder(FolderName: string; NewFolderName: string);
var
  FN: string;
  NFN: string;
  S: string;
  X: Integer;
  L: Integer;
begin
  FN := ToZipName(IncludeTrailingPathDelimiter(FolderName));
  NFN := ToZipName(IncludeTrailingPathDelimiter(NewFolderName));
  L := Length(FN);
  if Entries.IndexOf(NFN) = -1 then
  begin
    for X := 0 to Entries.Count - 1 do
    begin
      S := Entries.Items[X].FileName;
      if Pos(FN, S) = 1 then
      begin
        System.Delete(S, 1, L);
        S := NFN + S;
        Entries.Items[X].FileName := S;
        FIsDirty := True;
      end;
    end;
    if (FIsDirty) and (FBatchMode = False) then
      Rebuild;
  end;
end;

procedure TKAZip.SetReadOnly(const Value: Boolean);
begin
  FReadOnly := Value;
end;

procedure TKAZip.SetApplyAtributes(const Value: Boolean);
begin
  FApplyAttributes := Value;
end;

function TKAZip.AddEntryThroughStream(FileName: string; FileDate: TDateTime; FileAttr: Word): TStream;
var
  itemName: string;
  newEntry: TKAZIPEntriesEntry;
  i: Integer;
  zipComment: string;
  OBM: Boolean;
  newLocalEntryPosition: Int64;
  LFH: TLocalFileHeader;
begin
  if (FStoreFolders) and (FStoreRelativePath) then
    AddFolderChain(FileName);

  zipComment := Comment.Text;

  itemName := Filename;

  if not FStoreRelativePath then
    itemName := ExtractFileName(itemName);

  itemName := ToZipName(itemName);

  //If an item with this name already exists then remove it
  i := Entries.IndexOf(itemName);
  if i >= 0 then
  begin
    OBM := FBatchMode;
    try
      if OBM = False then
        FBatchMode := True;
      Remove(i);
    finally
      FBatchMode := OBM;
    end;
  end;

  //This is where the new local entry starts (where the central directly is).
  //We overwrite the central direct (and EOCD marker) and then re-write them after the end of the file
  newLocalEntryPosition := FEndOfCentralDir.OffsetOfStartOfCentralDirectory;

//  FCurrentDFS := uncompressedLength;

  //Fill records
  newEntry := TKAZipEntriesEntry(Entries.Add);

  //Local file entry
  newEntry.Date := FileDate;
  newEntry.CompressionMethod := ZipCompressionMethod_Deflate; //assume Deflate, but if the length is zero then it will be switched to Store (0)
  newEntry.CRC32 := 0; //don't know it yet, will back-fill
  newEntry.SizeCompressed := 0; //don't know it yet, will back-fill
  newEntry.SizeUncompressed := 0; //don't know it yet, will back-fill
  newEntry.FileName := ItemName;
  newEntry.Attributes := FileAttr;
  newEntry.LocalOffset := newLocalEntryPosition;

  // Write Local file entry to stream
  FZipStream.Position := newLocalEntryPosition;
  newEntry.FillLocalHeader(LFH);
  FZipStream.Write(LFH, SizeOf(LFH));
  if LFH.FilenameLength > 0 then
    FZipStream.Write(newEntry.FileName[1], LFH.FilenameLength);

  Result := TKAZipStream.Create(FZipStream, newEntry, FZipCompressionType, Self);
end;

function TKAZip.AddEntryThroughStream(FileName: string): TStream;
begin
  Result := AddEntryThroughStream(FileName, Now, faArchive);
end;

{ TCRC32Stream }

class function TCRC32Stream.CalcCRC32(const Data: AnsiString): LongWord;
var
  i: Integer;
begin
  //This is the old, and trusted, way. It doesn't let data stream
  Result := $FFFFFFFF;

  for i := 0 to Length(Data)-1 do
    Result := (Result shr 8) xor (CRCTable[Byte(Result) xor Ord(Data[i+1]) ]);

  Result := Result xor $FFFFFFFF;
end;

constructor TCRC32Stream.Create(TargetStream: TStream; StreamOwnership: TStreamOwnership=soReference);
begin
  inherited Create;

  FTargetStream := TargetStream;
  FTargetStreamOwnership := StreamOwnership;

  //Internal plumbing
  FCRC32 := $FFFFFFFF; //the initial state of a CRC32
  FTotalBytes := 0; 
end;

destructor TCRC32Stream.Destroy;
begin
  if FTargetStreamOwnership = soOwned then
    FTargetStream.Free;
  FTargetStream := nil;

  inherited;
end;

function TCRC32Stream.GetCRC32: LongWord;
begin
  Result := (FCRC32 xor $FFFFFFFF);
end;

function TCRC32Stream.Read(var Buffer; Count: Integer): Longint;
begin
  if FTargetStream = nil then
    raise Exception.Create('Cannot read from CRC stream when no target stream is present');

  Result := FTargetStream.Read({var}Buffer, Count)
end;

function TCRC32Stream.Seek(Offset: Integer; Origin: Word): Longint;
begin
  if FTargetStream = nil then
  begin
    //Seek returns the new Position property
    Result := 0;
    Exit;
  end;

  Result := FTargetStream.Seek(Offset, Origin);
end;

class procedure TCRC32Stream.SelfTest;

  procedure CheckEqualsString(const Expected, Actual: string);
  begin
    if Expected <> Actual then
      raise Exception.CreateFmt('Expected "%s", but was "%s"', [Expected, Actual]);
  end;

  procedure CheckEquals(const Expected, Actual: LongWord; Description: string);
  begin
    if Expected <> Actual then
      raise Exception.CreateFmt('Expected <%d>, but was <%d>.  %s', [Expected, Actual, Description]);
  end;

  function CRCNewWay(InputData: AnsiString): LongWord;
  var
    ss: TStringStream;
    cs: TCRC32Stream;
    crc1, crc2: LongWord;
    i: Integer;
  begin
    //process the inputData as a single chunk
    ss := TStringStream.Create('');
    try
      cs := TCRC32Stream.Create(ss, soReference);
      try
        cs.WriteBuffer(InputData[1], Length(InputData));
        crc1 := cs.CRC32;
      finally
        cs.Free;
      end;
      CheckEqualsString(InputData, ss.DataString);
    finally
      ss.Free;
    end;

    //process the input data byte-by-byte
    ss := TStringStream.Create('');
    try
      cs := TCRC32Stream.Create(ss, soReference);
      try
        for i := 1 to Length(InputData) do
          cs.WriteBuffer(InputData[i], 1);
        crc2 := cs.CRC32;
      finally
        cs.Free;
      end;
      CheckEqualsString(InputData, ss.DataString);
    finally
      ss.Free;
    end;

    //Check that both chunking methods match
    CheckEquals(crc1, crc2, InputData);

    //And return one of them to compare to the old way
    Result := crc1;
  end;

  procedure Test(InputString: AnsiString);
  var
    crcNew, crcOld: LongWord;
  begin
    crcOld := Self.CalcCRC32(InputString);
    crcNew := CRCNewWay(InputString);

    CheckEquals(crcOld, crcNew, InputString);
  end;

begin
  //http://www.asecuritysite.com/Encryption/crc32?word=Test%20vector%20from%20febooti.com
  Test(''); //00000000
  Test('1');
  Test('12');
  Test('123');
  Test('1234');
  Test('Test vector from febooti.com'); //0c877f61
  Test('The quick brown fox jumps over the lazy dog'); //0x414fa339 = 1095738169
  //A full buffer, followed by a partial fill, should lead to a full buffer again.
end;

function TCRC32Stream.Write(const Buffer; Count: Integer): Longint;
var
  i: Integer;
  state: LongWord;
  data: PByte;
begin
  state := FCRC32;

  data := Addr(Buffer);

  for i := 0 to Count-1 do
  begin
    state := (state shr 8) xor CRCTable[Byte(state) xor data^];
    Inc(data);
  end;

  FCRC32 := state;

  Inc(FTotalBytes, Count);

  if FTargetStream = nil then
  begin
    //Write returns the number of byts written
    Result := Count;
    Exit;
  end;

  Result := FTargetStream.Write(Buffer, Count)
end;

{ TKAZIPStream }

constructor TKAZIPStream.Create(TargetStream: TStream; Entry: TKAZipEntriesEntry; ZipCompressionType: TZipCompressionType; ParentZip: TKAZip);
var
  compressionLevel: TZCompressionLevel;
begin
  inherited Create;

  FTargetStream := TargetStream;
  FEntry := Entry;
  FParent := ParentZip;

  FStartPosition := FTargetStream.Position; //used to calculate size of compressed data

  //The zlib compression level to use
  case FParent.FZipCompressionType of
    ctNormal: compressionLevel := zcDefault;
    ctMaximum: compressionLevel := zcMax;
    ctFast: compressionLevel := zcFastest;
    ctSuperFast: compressionLevel := zcFastest;
    ctNone: compressionLevel := zcNone;
    else
      compressionLevel := zcDefault;
  end;

  {
    The standard ZLIB compresor does DEFLATE compression, but it adds a 2 byte header and 4 byte trailer
      header:
          CMF (Compression Method and flags)
          FLG (FLaGs)
      and final four bytes
          ADLER32 (Adler-32 checksum)
      There is a hack in zlib where if we specify a *negative* window size, it will use that to mean
      that we don't want it to add a header.
  }

  Fcompressor := TZCompressionStream.Create(FTargetStream, compressionLevel,
          -15, //windowBits - The default value is 15 if deflateInit is used instead.
          8, //memLevel - The default value is 8
          zsDefault); //strategy
//  OutputDebugString('SAMPLING ON');

  //A through-stream that calculates the crc32 of the uncompressed data streaming through it
  Fcrc32Stream := TCRC32Stream.Create(Fcompressor);
end;

destructor TKAZIPStream.Destroy;
var
  uncompressedSize: Int64;
  crc32: LongWord;
begin
  uncompressedSize := Fcrc32Stream.TotalBytesWritten;
  crc32 := FCrc32Stream.GetCRC32;
  FreeAndNil(FCrc32Stream);

  //Must free zlib compressor in order to flush it
  FreeAndNil(FCompressor);

  FinalizeEntryToZip(uncompressedSize, crc32);


  inherited;
end;

procedure TKAZIPStream.FinalizeEntryToZip(const uncompressedLength: Int64; const crc32: Longword);
var
  centralDirectoryPosition: Int64;
  compressedLength: Integer;
  compressionMethod: TZipCompressionMethod; //e.g. Store, Deflate
  //zipComment: string;
  LFH: TLocalFileHeader;
begin
  //The central directory will start here; just after the file we just wrote
  //Memorize it now, because we're going to backtrack and update the CRC and compressed size
  //also because we can use it to figure out the compressed size
  centralDirectoryPosition := FTargetStream.Position;

  compressedLength := (centralDirectoryPosition - FStartPosition);

  //If there was nothing compressed, then the compression method was "none", otherwise it is "deflate"
  if uncompressedLength > 0 then
    compressionMethod := ZipCompressionMethod_Deflate
  else
    compressionMethod := ZipCompressionMethod_Store;

  //Now that we know the CRC, the uncompressed, compressed sizes update them in the local file entry
  FEntry.Crc32 := crc32;
  FEntry.SizeUncompressed := uncompressedLength;
  FEntry.SizeCompressed := compressedLength;
  FEntry.CompressionMethod := compressionMethod;

  //Re-write the (now complete) local header
  try
    FTargetStream.Position := FEntry.LocalOffset;
    FEntry.FillLocalHeader(LFH);
    FTargetStream.Write(LFH, SizeOf(LFH));
  finally
    //Return to where the central directory will be, and write it
    FTargetStream.Position := centralDirectoryPosition;
  end;

  //Save the Central Directory entries
  FParent.RebuildCentralDirectory(FParent.ZipStream);
  {
  //Save EOCD (End of Central Directory) record
  FParent.FEndOfCentralDirPos := FParent.ZipStream.Position;
  FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory := centralDirectoryPosition;
  FParent.FEndOfCentralDir.SizeOfTheCentralDirectory := FParent.FEndOfCentralDirPos - FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
  Inc(FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
  Inc(FParent.FEndOfCentralDir.TotalNumberOfEntries);
  FParent.ZipStream.Write(FParent.FEndOfCentralDir, SizeOf(TEndOfCentralDir));

  //Save ZIP comment (if any)
  zipComment := FParent.Comment.Text;
  FParent.FZipCommentPos := FParent.ZipStream.Position;
  if Length(zipComment) > 0 then
    FParent.ZipStream.Write(zipComment[1], Length(zipComment));  }

  FParent.FIsDirty := True;
  if not FParent.FBatchMode then
  begin
    FParent.DoChange(FParent, ZIP_ITEM_ADDED);
  end;

end;

function TKAZIPStream.Read(var Buffer; Count: Integer): Longint;
begin
  raise Exception.Create('Cannot read from zip compressor stream');
end;

function TKAZIPStream.Seek(Offset: Integer; Origin: Word): Longint;
begin
  raise Exception.Create('Cannot seek a zip compressor stream');
end;

function TKAZIPStream.Write(const Buffer; Count: Integer): Longint;
begin
  Result := FCrc32Stream.Write(Buffer, Count);
end;

end.

