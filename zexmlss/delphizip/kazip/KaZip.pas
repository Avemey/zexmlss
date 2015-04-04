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


  {DoChange Events
    0 - Zip is Closed;
    1 - Zip is Opened;
    2 - Item is added to the zip
    3 - Item is removed from the Zip
    4 - Item comment changed
    5 - Item name changed
    6 - Item name changed
  }

  TLocalFile = packed record
    LocalFileHeaderSignature: Cardinal; //    4 bytes  = 0x04034b50 = 'PK'#03#04
    VersionNeededToExtract: WORD; //    2 bytes
    GeneralPurposeBitFlag: WORD; //    2 bytes
    CompressionMethod: TZipCompressionMethod; //    2 bytes
    LastModFileTimeDate: Cardinal; //    4 bytes
    Crc32: Cardinal; //    4 bytes
    CompressedSize: Cardinal; //    4 bytes
    UncompressedSize: Cardinal; //    4 bytes
    FilenameLength: WORD; //    2 bytes
    ExtraFieldLength: WORD; //    2 bytes
    //Data layout up to this point matches the actual ZIP format. Remaining fields are extra
    FileName: AnsiString; //    variable size
    ExtraField: AnsiString; //    variable size
    CompressedData: AnsiString; //    variable size
  end;

  TDataDescriptor = packed record
    DescriptorSignature: Cardinal; //    4 bytes UNDOCUMENTED
    Crc32: Cardinal; //    4 bytes
    CompressedSize: Cardinal; //    4 bytes
    UncompressedSize: Cardinal; //    4 bytes
  end;

  TCentralDirectoryFile = packed record
    CentralFileHeaderSignature: Cardinal; //    4 bytes = 0x02014b50 = 'PK'#01#02
    VersionMadeBy: WORD; //    2 bytes
    VersionNeededToExtract: WORD; //    2 bytes
    GeneralPurposeBitFlag: WORD; //    2 bytes
    CompressionMethod: TZipCompressionMethod; //    2 bytes
    LastModFileTimeDate: Cardinal; //    4 bytes
    Crc32: Cardinal; //    4 bytes
    CompressedSize: Cardinal; //    4 bytes
    UncompressedSize: Cardinal; //    4 bytes
    FilenameLength: WORD; //    2 bytes
    ExtraFieldLength: WORD; //    2 bytes
    FileCommentLength: WORD; //    2 bytes
    DiskNumberStart: WORD; //    2 bytes
    InternalFileAttributes: WORD; //    2 bytes
    ExternalFileAttributes: Cardinal; //    4 bytes
    RelativeOffsetOfLocalHeader: Cardinal; //    4 bytes
    FileName: AnsiString; //    variable size
    ExtraField: AnsiString; //    variable size
    FileComment: AnsiString; //    variable size
  end;

  TEndOfCentralDir = packed record
    EndOfCentralDirSignature: Cardinal; //    4 bytes = 0x06054b50 = 'PK'#05#06
    NumberOfThisDisk: WORD; //    2 bytes
    NumberOfTheDiskWithTheStart: WORD; //    2 bytes
    TotalNumberOfEntriesOnThisDisk: WORD; //    2 bytes
    TotalNumberOfEntries: WORD; //    2 bytes
    SizeOfTheCentralDirectory: Cardinal; //    4 bytes
    OffsetOfStartOfCentralDirectory: Cardinal; //    4 bytes
    ZipfileCommentLength: WORD; //    2 bytes
  end;

  TKAZipEntriesEntry = class(TCollectionItem)
  private
    FParent: TKAZipEntries;
    FCentralDirectoryFile: TCentralDirectoryFile;
    FLocalFile: TLocalFile;
    FIsEncrypted: Boolean;
    FIsFolder: Boolean;
    FDate: TDateTime;
    FCompressionType: TZipCompressionType;
    FSelected: Boolean;

    procedure SetSelected(const Value: Boolean);
    function GetLocalEntrySize: Cardinal;
    function GetCentralEntrySize: Cardinal;
    procedure SetComment(const Value: AnsiString);
    procedure SetFileName(const Value: AnsiString);
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    function GetCompressedData: AnsiString; overload;
    function GetCompressedData(Stream: TStream): Integer; overload;
    procedure ExtractToFile(FileName: string);
    procedure ExtractToStream(Stream: TStream);
    procedure SaveToFile(FileName: string);
    procedure SaveToStream(Stream: TStream);
    function Test: Boolean;

    property FileName: AnsiString read FCentralDirectoryFile.FileName write SetFileName;
    property Comment: AnsiString read FCentralDirectoryFile.FileComment write SetComment;
    property SizeUncompressed: Cardinal read FCentralDirectoryFile.UncompressedSize;
    property SizeCompressed: Cardinal read FCentralDirectoryFile.CompressedSize;
    property Date: TDateTime read FDate;
    property CRC32: Cardinal read FCentralDirectoryFile.CRC32;
    property Attributes: Cardinal read FCentralDirectoryFile.ExternalFileAttributes;
    property LocalOffset: Cardinal read FCentralDirectoryFile.RelativeOffsetOfLocalHeader;
    property IsEncrypted: Boolean read FIsEncrypted;
    property IsFolder: Boolean read FIsFolder;
    property BitFlag: Word read FCentralDirectoryFile.GeneralPurposeBitFlag;
    property CompressionMethod: TZipCompressionMethod read FCentralDirectoryFile.CompressionMethod;
    property CompressionType: TZipCompressionType read FCompressionType;
    property LocalEntrySize: Cardinal read GetLocalEntrySize;
    property CentralEntrySize: Cardinal read GetCentralEntrySize;
    property Selected: Boolean read FSelected write SetSelected;
  end;

  TKAZipEntries = class(TCollection)
  private
    { Private declarations }
    FParent: TKAZip;
    FIsZipFile: Boolean;
    FLocalHeaderNumFiles: Integer;

    function GetHeaderEntry(Index: Integer): TKAZipEntriesEntry;
    procedure SetHeaderEntry(Index: Integer; const Value: TKAZipEntriesEntry);
  protected
    { Protected declarations }
    function ReadBA(MS: TStream; Sz, Poz: Integer): TBytes;
    function Adler32(adler: uLong; buf: pByte; len: uInt): uLong;
    function CalcCRC32(const UncompressedData: string): Cardinal;
    function CalculateCRCFromStream(Stream: TStream): Cardinal;
    function RemoveRootName(const FileName, RootName: string): string;
    procedure SortList(List: TList);
    function FileTime2DateTime(FileTime: TFileTime): TDateTime;
    //**************************************************************************
    function FindCentralDirectory(MS: TStream): Boolean;
    function ParseCentralHeaders(MS: TStream): Boolean;
    function GetLocalEntry(MS: TStream; Offset: Integer; HeaderOnly: Boolean): TLocalFile;
    procedure LoadLocalHeaders(MS: TStream);
    function ParseLocalHeaders(MS: TStream): Boolean;

    //**************************************************************************
    procedure Remove(ItemIndex: Integer; Flush: Boolean); overload;
    procedure RemoveBatch(Files: TList);
    procedure InternalExtractToFile(Item: TKAZipEntriesEntry; FileName: string);
    //**************************************************************************
    function AddStreamFast(ItemName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry; overload;
    function AddStreamRebuild(ItemName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
    function AddFolderChain(ItemName: string): Boolean; overload;
    function AddFolderChain(ItemName: string; FileAttr: Word; FileDate: TDateTime): Boolean; overload;
    function AddFolderEx(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
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

    NewLHOffsets: array of Cardinal;
    NewEndOfCentralDir: TEndOfCentralDir;
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
    procedure LoadFromFile(FileName: string);
    procedure LoadFromStream(MS: TStream);
    //**************************************************************************
    procedure RebuildLocalFiles(MS: TStream);
    procedure RebuildCentralDirectory(MS: TStream);
    procedure RebuildEndOfCentralDirectory(MS: TStream);
    //**************************************************************************
    procedure OnDecompress(Sender: TObject);
    procedure OnCompress(Sender: TObject);
    procedure DoChange(Sender: TObject; const ChangeType: Integer); virtual;
    //**************************************************************************
    procedure WriteCentralDirectory(TargetStream: TStream);
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
    procedure ExtractToStream(Item: TKAZipEntriesEntry; Stream: TStream);
    procedure ExtractAll(TargetDirectory: string);
    procedure ExtractSelected(TargetDirectory: string);
    //**************************************************************************
    property Entries: TKAZipEntries read FZipHeader;
    property HasBadEntries: Boolean read FHasBadEntries;
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

  SIG_CentralFile =         $02014B50; // 'PK'#1#2
  SIG_LocalHeader =         $04034B50; // 'PK'#3#4
  SIG_EndOfCentralDirectory =  $06054B50; // 'PK'#5#6
  ZL_MULTIPLE_DISK_SIG =       $08074B50; // 'PK'#7#8
  ZL_DATA_DESCRIPT_SIG =       $08074B50; // 'PK'#7#8

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
  Default_TLocalFile: TLocalFile=();


procedure Register;
begin
  RegisterComponents('KA', [TKAZip]);
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

{ TKAZipEntriesEntry }

constructor TKAZipEntriesEntry.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FParent := TKAZipEntries(Collection);
  FSelected := False;
end;

destructor TKAZipEntriesEntry.Destroy;
begin

  inherited Destroy;
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
  BA: TLocalFile;
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

  BA := FParent.GetLocalEntry(FParent.FParent.FZipStream, LocalOffset, False);
  if BA.LocalFileHeaderSignature <> $04034B50 then
  begin
    Result := 0;
    Exit;
  end;
  if SizeCompressed > 0 then
    Result := Result + Stream.Write(BA.CompressedData[1], SizeCompressed);
end;

function TKAZipEntriesEntry.GetCompressedData: AnsiString;
var
  BA: TLocalFile;
  ZLHeader: TZLibStreamHeader;
  ZLH: Word;
  Compress: Byte;
begin
  Result := '';

  case CompressionMethod of
    ZipCompressionMethod_Store, ZipCompressionMethod_Deflate:
      begin
        BA := FParent.GetLocalEntry(FParent.FParent.FZipStream, LocalOffset, False);
        if BA.LocalFileHeaderSignature <> $04034B50 then
          Exit;
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
        Result := Result + BA.CompressedData;
      end;
  end;
end;

procedure TKAZipEntriesEntry.SetSelected(const Value: Boolean);
begin
  FSelected := Value;
end;

function TKAZipEntriesEntry.GetLocalEntrySize: Cardinal;
begin
  Result := SizeOf(TLocalFile) - 3 * SizeOf(string) +
      FCentralDirectoryFile.CompressedSize +
      FCentralDirectoryFile.FilenameLength +
      FCentralDirectoryFile.ExtraFieldLength;
  if (FCentralDirectoryFile.GeneralPurposeBitFlag and (1 shl 3)) > 0 then
  begin
    Result := Result + SizeOf(TDataDescriptor);
  end;
end;

function TKAZipEntriesEntry.GetCentralEntrySize: Cardinal;
begin
  Result := SizeOf(TCentralDirectoryFile) - 3 * SizeOf(string) +
    FCentralDirectoryFile.FilenameLength +
    FCentralDirectoryFile.ExtraFieldLength +
    FCentralDirectoryFile.FileCommentLength;
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
          Result := FParent.CalculateCRCFromStream(FS) = CRC32;
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
          Result := FParent.CalculateCRCFromStream(MS) = CRC32;
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
  FCentralDirectoryFile.FileComment := Value;
  FCentralDirectoryFile.FileCommentLength := Length(FCentralDirectoryFile.FileComment);
  FParent.Rebuild;
  if not FParent.FParent.FBatchMode then
  begin
    FParent.FParent.DoChange(FParent, 4);
  end;
end;

procedure TKAZipEntriesEntry.SetFileName(const Value: AnsiString);
var
  FN: string;
begin
  FN := ToZipName(Value);
  if FParent.IndexOf(FN) > -1 then
    raise Exception.Create('File with same name already exists in Archive!');
  FCentralDirectoryFile.FileName := ToZipName(Value);
  FCentralDirectoryFile.FilenameLength := Length(FCentralDirectoryFile.FileName);
  if not FParent.FParent.FBatchMode then
  begin
    FParent.Rebuild;
    FParent.FParent.DoChange(FParent, 5);
  end;
end;

{ TKAZipEntries }

constructor TKAZipEntries.Create(AOwner: TKAZip);
begin
  inherited Create(TKAZipEntriesEntry);
  FParent := AOwner;
  FIsZipFile := False;
end;

constructor TKAZipEntries.Create(AOwner: TKAZip; MS: TStream);
begin
  inherited Create(TKAZipEntriesEntry);
  FParent := AOwner;
  FIsZipFile := False;
  FLocalHeaderNumFiles := 0;
  ParseZip(MS);
end;

destructor TKAZipEntries.Destroy;
begin

  inherited Destroy;
end;

function TKAZipEntries.Adler32(adler: uLong; buf: pByte; len: uInt): uLong;
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

function TKAZipEntries.CalcCRC32(const UncompressedData: string): Cardinal;
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

function TKAZipEntries.CalculateCRCFromStream(Stream: TStream): Cardinal;
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

function TKAZipEntries.RemoveRootName(const FileName, RootName: string): string;
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

function TKAZipEntries.ReadBA(MS: TStream; Sz, Poz: Integer): TBytes;
begin
  SetLength(Result, SZ);
  MS.Position := Poz;
  MS.Read(Result[0], SZ);
end;

function TKAZipEntries.FindCentralDirectory(MS: TStream): Boolean;
var
  SeekStart: Integer;
  Poz: Integer;
  BR: Integer;
  Byte_: array[0..3] of Byte;
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
    MS.Read(Byte_, 4);
    if Byte_[0] = $50 then
    begin
      if (Byte_[1] = $4B)
        and (Byte_[2] = $05)
        and (Byte_[3] = $06) then
      begin
        MS.Position := Poz;
        FParent.FEndOfCentralDirPos := MS.Position;
        MS.Read(FParent.FEndOfCentralDir, SizeOf(FParent.FEndOfCentralDir));
        FParent.FZipCommentPos := MS.Position;
        FParent.FZipComment.Clear;
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

function TKAZipEntries.ParseCentralHeaders(MS: TStream): Boolean;
var
  X: Integer;
  Entry: TKAZipEntriesEntry;
  CDFile: TCentralDirectoryFile;
begin
  Result := False;
  try
    MS.Position := FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
    for X := 0 to FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk - 1 do
    begin
      //FillCh ar(CDFile, SizeOf(TCentralDirectoryFile), 0);  20140320: Filling a record containing managed types will cause a leak
      cdFile := Default_TCentralDirectoryFile;
      MS.Read(CDFile, SizeOf(TCentralDirectoryFile) - 3 * SizeOf(string));
      Entry := TKAZipEntriesEntry.Create(Self);
      Entry.FDate := FileDateToDateTime(CDFile.LastModFileTimeDate);
      if (CDFile.GeneralPurposeBitFlag and 1) > 0 then
        Entry.FIsEncrypted := True
      else
        Entry.FIsEncrypted := False;
      if CDFile.FilenameLength > 0 then
      begin
        SetLength(CDFile.FileName, CDFile.FilenameLength);
        MS.Read(CDFile.FileName[1], CDFile.FilenameLength)
      end;
      if CDFile.ExtraFieldLength > 0 then
      begin
        SetLength(CDFile.ExtraField, CDFile.ExtraFieldLength);
        MS.Read(CDFile.ExtraField[1], CDFile.ExtraFieldLength);
      end;
      if CDFile.FileCommentLength > 0 then
      begin
        SetLength(CDFile.FileComment, CDFile.FileCommentLength);
        MS.Read(CDFile.FileComment[1], CDFile.FileCommentLength);
      end;
      Entry.FIsFolder := (CDFile.ExternalFileAttributes and faDirectory) > 0;

      Entry.FCompressionType := ctUnknown;
      if (CDFile.CompressionMethod = ZipCompressionMethod_Deflate) or (CDFile.CompressionMethod = ZipCompressionMethod_Deflate64) then
      begin
        case CDFile.GeneralPurposeBitFlag and 6 of
          0: Entry.FCompressionType := ctNormal;
          2: Entry.FCompressionType := ctMaximum;
          4: Entry.FCompressionType := ctFast;
          6: Entry.FCompressionType := ctSuperFast
        end;
      end;
      Entry.FCentralDirectoryFile := CDFile;
      if Assigned(FParent.FOnZipOpen) then
        FParent.FOnZipOpen(FParent, X, FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
    end;
  except
    Exit;
  end;
  Result := Count = FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk;
end;

procedure TKAZipEntries.ParseZip(MS: TStream);
begin
  FIsZipFile := False;
  Clear;
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
      FIsZipFile := Count > 0;
      if FIsZipFile then
        FParent.FHasBadEntries := True;
    end;
  end;
end;

function TKAZipEntries.GetLocalEntry(MS: TStream; Offset: Integer; HeaderOnly: Boolean): TLocalFile;
var
  Byte_: array[0..4] of Byte;
  DataDescriptor: TDataDescriptor;
begin
  //20140320 -- You cannot ZeroMe mory/FillCh ar a structure containing strings.
  //It destroys some internal housekeeping and causes a leak.
//  FillCh ar(Result, SizeOf(Result), 0);
  Result := Default_TLocalFile;


  MS.Position := Offset;
  MS.Read(Byte_, 4);
  if not (
      (Byte_[0] = $50)
      and (Byte_[1] = $4B)
      and (Byte_[2] = $03)
      and (Byte_[3] = $04)) then
    Exit;

  MS.Position := Offset;
  MS.Read(Result, SizeOf(Result) - 3 * SizeOf(AnsiString));
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
  if (Result.GeneralPurposeBitFlag and (1 shl 3)) > 0 then
  begin
    MS.Read(DataDescriptor, SizeOf(TDataDescriptor));
    Result.Crc32 := DataDescriptor.Crc32;
    Result.CompressedSize := DataDescriptor.CompressedSize;
    Result.UnCompressedSize := DataDescriptor.UnCompressedSize;
  end;
  if not HeaderOnly then
  begin
    if Result.CompressedSize > 0 then
    begin
      SetLength(Result.CompressedData, Result.CompressedSize);
      MS.Read(Result.CompressedData[1], Result.CompressedSize);
    end;
  end;
end;

procedure TKAZipEntries.LoadLocalHeaders(MS: TStream);
var
  X: Integer;
begin
  FParent.FHasBadEntries := False;
  for X := 0 to Count - 1 do
  begin
    if Assigned(FParent.FOnZipOpen) then
      FParent.FOnZipOpen(FParent, X, FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
    Items[X].FLocalFile := GetLocalEntry(MS, Items[X].FCentralDirectoryFile.RelativeOffsetOfLocalHeader, True);
    if Items[X].FLocalFile.LocalFileHeaderSignature <> $04034B50 then
      FParent.FHasBadEntries := True;
  end;
end;

function TKAZipEntries.ParseLocalHeaders(MS: TStream): Boolean;
var
  Poz: Integer;
  NLE: Integer;
  Byte_: array[0..4] of Byte;
  LocalFile: TLocalFile;
  DataDescriptor: TDataDescriptor;
  Entry: TKAZipEntriesEntry;
  CDFile: TCentralDirectoryFile;
  CDSize: Cardinal;
  L: Integer;
  NoMore: Boolean;
begin
  Result := False;
  FLocalHeaderNumFiles := 0;
  Clear;
  try
    Poz := 0;
    NLE := 0;
    CDSize := 0;
    repeat
      NoMore := True;
      MS.Position := Poz;
      MS.Read(Byte_, 4);
      if (Byte_[0] = $50)
        and (Byte_[1] = $4B)
        and (Byte_[2] = $03)
        and (Byte_[3] = $04) then
      begin
        Result := True;
        Inc(FLocalHeaderNumFiles);
        NoMore := False;
        MS.Position := Poz;
        MS.Read(LocalFile, SizeOf(TLocalFile) - 3 * SizeOf(string));
        if LocalFile.FilenameLength > 0 then
        begin
          SetLength(LocalFile.FileName, LocalFile.FilenameLength);
          MS.Read(LocalFile.FileName[1], LocalFile.FilenameLength);
        end;
        if LocalFile.ExtraFieldLength > 0 then
        begin
          SetLength(LocalFile.ExtraField, LocalFile.ExtraFieldLength);
          MS.Read(LocalFile.ExtraField[1], LocalFile.ExtraFieldLength);
        end;
        if (LocalFile.GeneralPurposeBitFlag and (1 shl 3)) > 0 then
        begin
          MS.Read(DataDescriptor, SizeOf(TDataDescriptor));
          LocalFile.Crc32 := DataDescriptor.Crc32;
          LocalFile.CompressedSize := DataDescriptor.CompressedSize;
          LocalFile.UncompressedSize := DataDescriptor.UncompressedSize;
        end;
        MS.Position := MS.Position + LongInt(LocalFile.CompressedSize);

        //FillCha r(CDFile, SizeOf(TCentralDirectoryFile), 0);  20140320 Cannot do FillCh ar on a structure containing managed types
        CDFile := Default_TCentralDirectoryFile;
        CDFile.CentralFileHeaderSignature := SIG_CentralFile;
        CDFile.VersionMadeBy := 20;
        CDFile.VersionNeededToExtract := LocalFile.VersionNeededToExtract;
        CDFile.GeneralPurposeBitFlag := LocalFile.GeneralPurposeBitFlag;
        CDFile.CompressionMethod := LocalFile.CompressionMethod;
        CDFile.LastModFileTimeDate := LocalFile.LastModFileTimeDate;
        CDFile.Crc32 := LocalFile.Crc32;
        CDFile.CompressedSize := LocalFile.CompressedSize;
        CDFile.UncompressedSize := LocalFile.UncompressedSize;
        CDFile.FilenameLength := LocalFile.FilenameLength;
        CDFile.ExtraFieldLength := LocalFile.ExtraFieldLength;
        CDFile.FileCommentLength := 0;
        CDFile.DiskNumberStart := 0;
        CDFile.InternalFileAttributes := LocalFile.VersionNeededToExtract;
        CDFile.ExternalFileAttributes := faArchive;
        CDFile.RelativeOffsetOfLocalHeader := Poz;
        CDFile.FileName := LocalFile.FileName;
        L := Length(CDFile.FileName);
        if L > 0 then
        begin
          if CDFile.FileName[L] = '/' then
            CDFile.ExternalFileAttributes := faDirectory;
        end;
        CDFile.ExtraField := LocalFile.ExtraField;
        CDFile.FileComment := '';

        Entry := TKAZipEntriesEntry.Create(Self);
        Entry.FDate := FileDateToDateTime(CDFile.LastModFileTimeDate);
        if (CDFile.GeneralPurposeBitFlag and 1) > 0 then
          Entry.FIsEncrypted := True
        else
          Entry.FIsEncrypted := False;
        Entry.FIsFolder := (CDFile.ExternalFileAttributes and faDirectory) > 0;
        Entry.FCompressionType := ctUnknown;
        if (CDFile.CompressionMethod = ZipCompressionMethod_Deflate) or (CDFile.CompressionMethod = ZipCompressionMethod_Deflate64) then
        begin
          case CDFile.GeneralPurposeBitFlag and 6 of
            0: Entry.FCompressionType := ctNormal;
            2: Entry.FCompressionType := ctMaximum;
            4: Entry.FCompressionType := ctFast;
            6: Entry.FCompressionType := ctSuperFast
          end;
        end;
        Entry.FCentralDirectoryFile := CDFile;
        Poz := MS.Position;
        Inc(NLE);
        CDSize := CDSize + Entry.CentralEntrySize;
      end;
    until NoMore;

    FParent.FEndOfCentralDir.EndOfCentralDirSignature := SIG_EndOfCentralDirectory; //PK 0x05 0x06
    FParent.FEndOfCentralDir.NumberOfThisDisk := 0;
    FParent.FEndOfCentralDir.NumberOfTheDiskWithTheStart := 0;
    FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk := NLE;
    FParent.FEndOfCentralDir.SizeOfTheCentralDirectory := CDSize;
    FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory := MS.Position;
    FParent.FEndOfCentralDir.ZipfileCommentLength := 0;
  except
    Exit;
  end;
end;

procedure TKAZipEntries.Remove(ItemIndex: Integer; Flush: Boolean);
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
begin
  TargetPos := Items[ItemIndex].FCentralDirectoryFile.RelativeOffsetOfLocalHeader;
  ShiftSize := Items[ItemIndex].LocalEntrySize;
//  BufStart := TargetPos + ShiftSize;
//  BufLen := FParent.FZipStream.Size - BufStart;
  Border := TargetPos;
  Delete(ItemIndex);
  if (FParent.FZipSaveMethod = FastSave) and (Count > 0) then
  begin
    ZipComment := FParent.Comment.Text;

{
    SetLength(BUF,BufLen);
    FParent.FZipStream.Position := BufStart;
    NR := FParent.FZipStream.Read(BUF[1],BufLen);

    FParent.FZipStream.Position := TargetPos;
    NW := FParent.FZipStream.Write(BUF[1],BufLen);
    SetLength(BUF,0);
  }

    for X := 0 to Count - 1 do
    begin
      if Items[X].FCentralDirectoryFile.RelativeOffsetOfLocalHeader > Border then
      begin
        Dec(Items[X].FCentralDirectoryFile.RelativeOffsetOfLocalHeader, ShiftSize);
        TargetPos := TargetPos + Items[X].LocalEntrySize;
      end
    end;

    FParent.FZipStream.Position := TargetPos;
    //************************************ MARK START OF CENTRAL DIRECTORY
    FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory := FParent.FZipStream.Position;
    //************************************ SAVE CENTRAL DIRECTORY
    for X := 0 to Count - 1 do
    begin
      FParent.FZipStream.Write(Self.Items[X].FCentralDirectoryFile, SizeOf(Self.Items[X].FCentralDirectoryFile) - 3 * SizeOf(string));
      if Self.Items[X].FCentralDirectoryFile.FilenameLength > 0 then
        FParent.FZipStream.Write(Self.Items[X].FCentralDirectoryFile.FileName[1], Self.Items[X].FCentralDirectoryFile.FilenameLength);
      if Self.Items[X].FCentralDirectoryFile.ExtraFieldLength > 0 then
        FParent.FZipStream.Write(Self.Items[X].FCentralDirectoryFile.ExtraField[1], Self.Items[X].FCentralDirectoryFile.ExtraFieldLength);
      if Self.Items[X].FCentralDirectoryFile.FileCommentLength > 0 then
        FParent.FZipStream.Write(Self.Items[X].FCentralDirectoryFile.FileComment[1], Self.Items[X].FCentralDirectoryFile.FileCommentLength);
    end;
    //************************************ SAVE END CENTRAL DIRECTORY RECORD
    FParent.FEndOfCentralDirPos := FParent.FZipStream.Position;
    FParent.FEndOfCentralDir.SizeOfTheCentralDirectory := FParent.FEndOfCentralDirPos - FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
    Dec(FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
    Dec(FParent.FEndOfCentralDir.TotalNumberOfEntries);
    FParent.FZipStream.Write(FParent.FEndOfCentralDir, SizeOf(TEndOfCentralDir));
    //************************************ SAVE ZIP COMMENT IF ANY
    FParent.FZipCommentPos := FParent.FZipStream.Position;
    if Length(ZipComment) > 0 then
    begin
      FParent.FZipStream.Write(ZipComment[1], Length(ZipComment));
    end;
    FParent.FZipStream.Size := FParent.FZipStream.Position;
  end
  else
  begin
    if FParent.FUseTempFiles then
    begin
      TempFileName := FParent.GetDelphiTempFileName;
      TempStream := TFileStream.Create(TempFileName, fmOpenReadWrite or FmCreate);
      try
        FParent.SaveToStream(TempStream);
        TempStream.Position := 0;
        OSL := FParent.FZipStream.Size;
        try
          FParent.FZipStream.Size := TempStream.Size;
        except
          FParent.FZipStream.Size := OSL;
          raise;
        end;
        FParent.FZipStream.Position := 0;
        FParent.FZipStream.CopyFrom(TempStream, TempStream.Size);
        //*********************************************************************
        FParent.FZipHeader.ParseZip(FParent.FZipStream);
        //*********************************************************************
      finally
        TempStream.Free;
        DeleteFile(TempFileName)
      end;
    end
    else
    begin
      NewSize := 0;
      for X := 0 to Count - 1 do
      begin
        NewSize := NewSize + Items[X].LocalEntrySize + Items[X].CentralEntrySize;
        if Assigned(FParent.FOnRemoveItems) then
          FParent.FOnRemoveItems(FParent, X, Count - 1);
      end;
      NewSize := NewSize + SizeOf(FParent.FEndOfCentralDir) + FParent.FEndOfCentralDir.ZipfileCommentLength;
      TempMSStream := TMemoryStream.Create;
      try
        TempMSStream.SetSize(NewSize);
        TempMSStream.Position := 0;
        FParent.SaveToStream(TempMSStream);
        TempMSStream.Position := 0;
        OSL := FParent.FZipStream.Size;
        try
          FParent.FZipStream.Size := TempMSStream.Size;
        except
          FParent.FZipStream.Size := OSL;
          raise;
        end;
        FParent.FZipStream.Position := 0;
        FParent.FZipStream.CopyFrom(TempMSStream, TempMSStream.Size);
        //*********************************************************************
        FParent.FZipHeader.ParseZip(FParent.FZipStream);
        //*********************************************************************
      finally
        TempMSStream.Free;
      end;
    end;
  end;
  FParent.FIsDirty := True;
  if not FParent.FBatchMode then
  begin
    FParent.DoChange(FParent, 3);
  end;
end;

procedure TKAZipEntries.Remove(ItemIndex: Integer);
begin
  Remove(ItemIndex, True);
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

procedure TKAZipEntries.RemoveBatch(Files: TList);
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
    Delete(Integer(Files.Items[X]));
    if Assigned(FParent.FOnRemoveItems) then
      FParent.FOnRemoveItems(FParent, Files.Count - X, Files.Count);
  end;
  NewSize := 0;
  if FParent.FUseTempFiles then
  begin
    TempFileName := FParent.GetDelphiTempFileName;
    TempStream := TFileStream.Create(TempFileName, fmOpenReadWrite or FmCreate);
    try
      FParent.SaveToStream(TempStream);
      TempStream.Position := 0;
      OSL := FParent.FZipStream.Size;
      try
        FParent.FZipStream.Size := TempStream.Size;
      except
        FParent.FZipStream.Size := OSL;
        raise;
      end;
      FParent.FZipStream.Position := 0;
      FParent.FZipStream.CopyFrom(TempStream, TempStream.Size);
      //*********************************************************************
      FParent.FZipHeader.ParseZip(FParent.FZipStream);
      //*********************************************************************
    finally
      TempStream.Free;
      DeleteFile(TempFileName)
    end;
  end
  else
  begin
    for X := 0 to Count - 1 do
    begin
      NewSize := NewSize + Items[X].LocalEntrySize + Items[X].CentralEntrySize;
      if Assigned(FParent.FOnRemoveItems) then
        FParent.FOnRemoveItems(FParent, X, Count - 1);
    end;
    NewSize := NewSize + SizeOf(FParent.FEndOfCentralDir) + FParent.FEndOfCentralDir.ZipfileCommentLength;
    TempMSStream := TMemoryStream.Create;
    try
      TempMSStream.SetSize(NewSize);
      TempMSStream.Position := 0;
      FParent.SaveToStream(TempMSStream);
      TempMSStream.Position := 0;
      OSL := FParent.FZipStream.Size;
      try
        FParent.FZipStream.Size := TempMSStream.Size;
      except
        FParent.FZipStream.Size := OSL;
        raise;
      end;
      FParent.FZipStream.Position := 0;
      FParent.FZipStream.CopyFrom(TempMSStream, TempMSStream.Size);
      //*********************************************************************
      FParent.FZipHeader.ParseZip(FParent.FZipStream);
      //*********************************************************************
    finally
      TempMSStream.Free;
    end;
  end;
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
    if AnsiCompareText(FN, ToZipName(Items[X].FCentralDirectoryFile.FileName)) = 0 then
    begin
      Result := X;
      Exit;
    end;
  end;
end;

function TKAZipEntries.AddStreamFast(ItemName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
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
begin
  //*********************************** COMPRESS DATA
  zipComment := FParent.Comment.Text;

  if not FParent.FStoreRelativePath then
    ItemName := ExtractFileName(ItemName);

  ItemName := ToZipName(ItemName);

  //If an item with this name already exists then remove it
  i := Self.IndexOf(ItemName);
  if i >= 0 then
  begin
    OBM := FParent.FBatchMode;
    try
      if OBM = False then
        FParent.FBatchMode := True;
      Remove(i);
    finally
      FParent.FBatchMode := OBM;
    end;
  end;

  //This is where the new local entry starts (where the central directly is).
  //We overwrite the central direct (and EOCD marker) and then re-write them after the end of the file
  newLocalEntryPosition := FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory;

  uncompressedLength := Stream.Size - Stream.Position;
  FParent.FCurrentDFS := uncompressedLength;


  dataCrc32 := 0; //we don't know it yet; back-fill it
//  compressedLength := 0; //we don't know it yet; back-fill it

  if uncompressedLength > 0 then
    compressionMode := ZipCompressionMethod_Deflate
  else
    compressionMode := ZipCompressionMethod_Store;

  //Fill records
  Result := TKAZipEntriesEntry(Self.Add);

  //Local file entry
  Result.FLocalFile.LocalFileHeaderSignature := SIG_LocalHeader; //'PK'#03#04
  Result.FLocalFile.VersionNeededToExtract := 20;
  Result.FLocalFile.GeneralPurposeBitFlag := 0;
  Result.FLocalFile.CompressionMethod := compressionMode;
  Result.FLocalFile.LastModFileTimeDate := DateTimeToFileDate(FileDate);
  Result.FLocalFile.Crc32 := 0; //don't know it yet, will back-fill
  Result.FLocalFile.CompressedSize := 0; //don't know it yet, will back-fill
  Result.FLocalFile.UncompressedSize := uncompressedLength;
  Result.FLocalFile.FilenameLength := Length(ItemName);
  Result.FLocalFile.ExtraFieldLength := 0;
  Result.FLocalFile.FileName := ItemName;
  Result.FLocalFile.ExtraField := '';
  Result.FLocalFile.CompressedData := ''; //not used

  // Write Local file entry to stream
  FParent.FZipStream.Position := newLocalEntryPosition;
  FParent.FZipStream.Write(Result.FLocalFile, SizeOf(Result.FLocalFile) - 3 * SizeOf(string)); //excluding FileName, ExtraField, and CompressedData
  if Result.FLocalFile.FilenameLength > 0 then
    FParent.FZipStream.Write(Result.FLocalFile.FileName[1], Result.FLocalFile.FilenameLength);

  //Now stream the compressed data
  startOfCompressedDataPosition := FParent.FZipStream.Position;

  if compressionMode = ZipCompressionMethod_Deflate then
  begin
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
    compressor := TZCompressionStream.Create(FParent.FZipStream, compressionLevel,
          -15, //windowBits - The default value is 15 if deflateInit is used instead.
          8, //memLevel - The default value is 8
          zsDefault); //strategy
    try
      compressor.OnProgress := FParent.OnCompress;

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
  centralDirectoryPosition := FParent.FZipStream.Position;
  try
    compressedLength := (centralDirectoryPosition - startOfCompressedDataPosition);

    //Now that we know the CRC and the compressed size, update them in the local file entry
    Result.FLocalFile.Crc32 := dataCrc32;
    Result.FLocalFile.CompressedSize := compressedLength;
    FParent.FZipStream.Position := newLocalEntryPosition;
    FParent.FZipStream.Write(Result.FLocalFile, SizeOf(Result.FLocalFile) - 3 * SizeOf(string)); //excluding FileName, ExtraField, and CompressedData
  finally
    //Return to where the central directory will be, and write it
    FParent.FZipStream.Position := centralDirectoryPosition;
  end;

  //Create the Central Directory entry
  Result.FCentralDirectoryFile.CentralFileHeaderSignature := SIG_CentralFile; //PK 0x01 0x02
  Result.FCentralDirectoryFile.VersionMadeBy := 20;
  Result.FCentralDirectoryFile.VersionNeededToExtract := 20;
  Result.FCentralDirectoryFile.GeneralPurposeBitFlag := 0;
  Result.FCentralDirectoryFile.CompressionMethod := compressionMode;
  Result.FCentralDirectoryFile.LastModFileTimeDate := DateTimeToFileDate(FileDate);
  Result.FCentralDirectoryFile.Crc32 := dataCrc32;
  Result.FCentralDirectoryFile.CompressedSize := compressedLength;
  Result.FCentralDirectoryFile.UncompressedSize := uncompressedLength;
  Result.FCentralDirectoryFile.FilenameLength := Length(ItemName);
  Result.FCentralDirectoryFile.ExtraFieldLength := 0;
  Result.FCentralDirectoryFile.FileCommentLength := 0;
  Result.FCentralDirectoryFile.DiskNumberStart := 0;
  Result.FCentralDirectoryFile.InternalFileAttributes := 0;
  Result.FCentralDirectoryFile.ExternalFileAttributes := FileAttr;
  Result.FCentralDirectoryFile.RelativeOffsetOfLocalHeader := newLocalEntryPosition;
  Result.FCentralDirectoryFile.FileName := ItemName;
  Result.FCentralDirectoryFile.ExtraField := '';
  Result.FCentralDirectoryFile.FileComment := '';

  //Save the Central Directory entries (todo: split this into separate function)
  for i := 0 to Self.Count-1 do
  begin
    //packed byte data
    FParent.FZipStream.Write(Self.Items[i].FCentralDirectoryFile, SizeOf(Self.Items[i].FCentralDirectoryFile) - 3 * SizeOf(string));
    //optional filenmae
    if Self.Items[i].FCentralDirectoryFile.FilenameLength > 0 then
      FParent.FZipStream.Write(Self.Items[i].FCentralDirectoryFile.FileName[1], Self.Items[i].FCentralDirectoryFile.FilenameLength);
    //optional extra field
    if Self.Items[i].FCentralDirectoryFile.ExtraFieldLength > 0 then
      FParent.FZipStream.Write(Self.Items[i].FCentralDirectoryFile.ExtraField[1], Self.Items[i].FCentralDirectoryFile.ExtraFieldLength);
    //optional file comment
    if Self.Items[i].FCentralDirectoryFile.FileCommentLength > 0 then
      FParent.FZipStream.Write(Self.Items[i].FCentralDirectoryFile.FileComment[1], Self.Items[i].FCentralDirectoryFile.FileCommentLength);
  end;

  //Save EOCD (End of Central Directory) record
  FParent.FEndOfCentralDirPos := FParent.FZipStream.Position;
  FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory := centralDirectoryPosition;
  FParent.FEndOfCentralDir.SizeOfTheCentralDirectory := FParent.FEndOfCentralDirPos - FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
  Inc(FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
  Inc(FParent.FEndOfCentralDir.TotalNumberOfEntries);
  FParent.FZipStream.Write(FParent.FEndOfCentralDir, SizeOf(TEndOfCentralDir));

  //Save ZIP comment (if any)
  FParent.FZipCommentPos := FParent.FZipStream.Position;
  if Length(zipComment) > 0 then
    FParent.FZipStream.Write(zipComment[1], Length(zipComment));

  Result.FDate := FileDate;

  if (Result.FCentralDirectoryFile.GeneralPurposeBitFlag and 1) > 0 then
    Result.FIsEncrypted := True
  else
    Result.FIsEncrypted := False;
  Result.FIsFolder := (Result.FCentralDirectoryFile.ExternalFileAttributes and faDirectory) > 0;
  Result.FCompressionType := ctUnknown;
  if (Result.FCentralDirectoryFile.CompressionMethod = ZipCompressionMethod_Deflate) or (Result.FCentralDirectoryFile.CompressionMethod = ZipCompressionMethod_Deflate64) then
  begin
    case Result.FCentralDirectoryFile.GeneralPurposeBitFlag and 6 of
      0: Result.FCompressionType := ctNormal;
      2: Result.FCompressionType := ctMaximum;
      4: Result.FCompressionType := ctFast;
      6: Result.FCompressionType := ctSuperFast
    end;
  end;
  FParent.FIsDirty := True;
  if not FParent.FBatchMode then
  begin
    FParent.DoChange(FParent, 2);
  end;
end;

function TKAZipEntries.AddStreamRebuild(ItemName: string;
  FileAttr: Word;
  FileDate: TDateTime;
  Stream: TStream): TKAZipEntriesEntry;
var
  Compressor: TZCompressionStream;
  CS: TStringStream;
  cm: Word; //CompressionMethod to use
  S: string;
  UL: Integer;
  CL: Integer;
  I: Integer;
  X: Integer;
  FCRC32: Cardinal;
  OSL: Cardinal;
  NewSize: Cardinal;
  ZipComment: string;
  TempStream: TFileStream;
  TempMSStream: TMemoryStream;
  TempFileName: string;
  Level: TZCompressionLevel;
  OBM: Boolean;
begin
  if FParent.FUseTempFiles then
  begin
    TempFileName := FParent.GetDelphiTempFileName;
    TempStream := TFileStream.Create(TempFileName, fmOpenReadWrite or FmCreate);
    try
      //*********************************** SAVE ALL OLD LOCAL ITEMS
      FParent.RebuildLocalFiles(TempStream);
      //*********************************** COMPRESS DATA
      ZipComment := FParent.Comment.Text;
      if not FParent.FStoreRelativePath then
        ItemName := ExtractFileName(ItemName);
      ItemName := ToZipName(ItemName);
      I := IndexOf(ItemName);
      if I > -1 then
      begin
        OBM := FParent.FBatchMode;
        try
          if OBM = False then
            FParent.FBatchMode := True;
          Remove(I);
        finally
          FParent.FBatchMode := OBM;
        end;
      end;

      cm := 0;
      CS := TStringStream.Create('');
      CS.Position := 0;
      try
        UL := Stream.Size - Stream.Position;
        SetLength(S, UL);
        if UL > 0 then
        begin
          Stream.Read(S[1], UL);
          cm := ZipCompressionMethod_Deflate;
        end;
        FCRC32 := CalcCRC32(S);
        FParent.FCurrentDFS := UL;

        case FParent.FZipCompressionType of
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
            Compressor.OnProgress := FParent.OnCompress;
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
      Result := TKAZipEntriesEntry(Self.Add);
      with Result.FLocalFile do
      begin
        LocalFileHeaderSignature := SIG_LocalHeader;
        VersionNeededToExtract := 20;
        GeneralPurposeBitFlag := 0;
        CompressionMethod := cm;
        LastModFileTimeDate := DateTimeToFileDate(FileDate);
        Crc32 := FCRC32;
        CompressedSize := CL;
        UncompressedSize := UL;
        FilenameLength := Length(ItemName);
        ExtraFieldLength := 0;
        FileName := ItemName;
        ExtraField := '';
        CompressedData := '';
      end;

      with Result.FCentralDirectoryFile do
      begin
        CentralFileHeaderSignature := SIG_CentralFile;
        VersionMadeBy := 20;
        VersionNeededToExtract := 20;
        GeneralPurposeBitFlag := 0;
        CompressionMethod := CM;
        LastModFileTimeDate := DateTimeToFileDate(FileDate);
        Crc32 := FCRC32;
        CompressedSize := CL;
        UncompressedSize := UL;
        FilenameLength := Length(ItemName);
        ExtraFieldLength := 0;
        FileCommentLength := 0;
        DiskNumberStart := 0;
        InternalFileAttributes := 0;
        ExternalFileAttributes := FileAttr;
        RelativeOffsetOfLocalHeader := TempStream.Position;
        FileName := ItemName;
        ExtraField := '';
        FileComment := '';
      end;

      //************************************ SAVE LOCAL HEADER AND COMPRESSED DATA
      TempStream.Write(Result.FLocalFile, SizeOf(Result.FLocalFile) - 3 * SizeOf(string));
      if Result.FLocalFile.FilenameLength > 0 then
        TempStream.Write(Result.FLocalFile.FileName[1], Result.FLocalFile.FilenameLength);
      if CL > 0 then
        TempStream.Write(S[1], CL);
      //************************************
      FParent.NewLHOffsets[Count - 1] := Result.FCentralDirectoryFile.RelativeOffsetOfLocalHeader;
      FParent.RebuildCentralDirectory(TempStream);
      FParent.RebuildEndOfCentralDirectory(TempStream);
      //************************************
      TempStream.Position := 0;
      OSL := FParent.FZipStream.Size;
      try
        FParent.FZipStream.Size := TempStream.Size;
      except
        FParent.FZipStream.Size := OSL;
        raise;
      end;
      FParent.FZipStream.Position := 0;
      FParent.FZipStream.CopyFrom(TempStream, TempStream.Size);
    finally
      TempStream.Free;
      DeleteFile(TempFileName)
    end;
  end
  else
  begin
    TempMSStream := TMemoryStream.Create;
    NewSize := 0;
    for X := 0 to Count - 1 do
    begin
      NewSize := NewSize + Items[X].LocalEntrySize + Items[X].CentralEntrySize;
      if Assigned(FParent.FOnRemoveItems) then
        FParent.FOnRemoveItems(FParent, X, Count - 1);
    end;
    NewSize := NewSize + SizeOf(FParent.FEndOfCentralDir) + FParent.FEndOfCentralDir.ZipfileCommentLength;
    try
      TempMSStream.SetSize(NewSize);
      TempMSStream.Position := 0;
      //*********************************** SAVE ALL OLD LOCAL ITEMS
      FParent.RebuildLocalFiles(TempMSStream);
      //*********************************** COMPRESS DATA
      ZipComment := FParent.Comment.Text;
      if not FParent.FStoreRelativePath then
        ItemName := ExtractFileName(ItemName);
      ItemName := ToZipName(ItemName);
      I := IndexOf(ItemName);
      if I > -1 then
      begin
        OBM := FParent.FBatchMode;
        try
          if OBM = False then
            FParent.FBatchMode := True;
          Remove(I);
        finally
          FParent.FBatchMode := OBM;
        end;
      end;

      CM := 0;
      CS := TStringStream.Create('');
      CS.Position := 0;
      try
        UL := Stream.Size - Stream.Position;
        SetLength(S, UL);
        if UL > 0 then
        begin
          Stream.Read(S[1], UL);
          CM := ZipCompressionMethod_Deflate;
        end;
        FCRC32 := CalcCRC32(S);
        FParent.FCurrentDFS := UL;

        case FParent.FZipCompressionType of
          ctNormal: Level := zcDefault;
          ctMaximum: Level := zcMax;
          ctFast: Level := zcFastest;
          ctSuperFast: Level := zcFastest;
          ctNone: Level := zcNone;
          else
            Level := zcDefault;
        end;

        if CM = ZipCompressionMethod_Deflate then
        begin
          Compressor := TZCompressionStream.Create(CS, Level);
          try
            Compressor.OnProgress := FParent.OnCompress;
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
      Result := TKAZipEntriesEntry(Self.Add);
      with Result.FLocalFile do
      begin
        LocalFileHeaderSignature := SIG_LocalHeader;
        VersionNeededToExtract := 20;
        GeneralPurposeBitFlag := 0;
        CompressionMethod := CM;
        LastModFileTimeDate := DateTimeToFileDate(FileDate);
        Crc32 := FCRC32;
        CompressedSize := CL;
        UncompressedSize := UL;
        FilenameLength := Length(ItemName);
        ExtraFieldLength := 0;
        FileName := ItemName;
        ExtraField := '';
        CompressedData := '';
      end;

      with Result.FCentralDirectoryFile do
      begin
        CentralFileHeaderSignature := SIG_CentralFile;
        VersionMadeBy := 20;
        VersionNeededToExtract := 20;
        GeneralPurposeBitFlag := 0;
        CompressionMethod := CM;
        LastModFileTimeDate := DateTimeToFileDate(FileDate);
        Crc32 := FCRC32;
        CompressedSize := CL;
        UncompressedSize := UL;
        FilenameLength := Length(ItemName);
        ExtraFieldLength := 0;
        FileCommentLength := 0;
        DiskNumberStart := 0;
        InternalFileAttributes := 0;
        ExternalFileAttributes := FileAttr;
        RelativeOffsetOfLocalHeader := TempMSStream.Position;
        FileName := ItemName;
        ExtraField := '';
        FileComment := '';
      end;

      //************************************ SAVE LOCAL HEADER AND COMPRESSED DATA
      TempMSStream.Write(Result.FLocalFile, SizeOf(Result.FLocalFile) - 3 * SizeOf(string));
      if Result.FLocalFile.FilenameLength > 0 then
        TempMSStream.Write(Result.FLocalFile.FileName[1], Result.FLocalFile.FilenameLength);
      if CL > 0 then
        TempMSStream.Write(S[1], CL);
      //************************************
      FParent.NewLHOffsets[Count - 1] := Result.FCentralDirectoryFile.RelativeOffsetOfLocalHeader;
      FParent.RebuildCentralDirectory(TempMSStream);
      FParent.RebuildEndOfCentralDirectory(TempMSStream);
      //************************************
      TempMSStream.Position := 0;
      OSL := FParent.FZipStream.Size;
      try
        FParent.FZipStream.Size := TempMSStream.Size;
      except
        FParent.FZipStream.Size := OSL;
        raise;
      end;
      FParent.FZipStream.Position := 0;
      FParent.FZipStream.CopyFrom(TempMSStream, TempMSStream.Size);
    finally
      TempMSStream.Free;
    end;
  end;

  Result.FDate := FileDateToDateTime(Result.FCentralDirectoryFile.LastModFileTimeDate);
  if (Result.FCentralDirectoryFile.GeneralPurposeBitFlag and 1) > 0 then
    Result.FIsEncrypted := True
  else
    Result.FIsEncrypted := False;
  Result.FIsFolder := (Result.FCentralDirectoryFile.ExternalFileAttributes and faDirectory) > 0;
  Result.FCompressionType := ctUnknown;
  if (Result.FCentralDirectoryFile.CompressionMethod = ZipCompressionMethod_Deflate) or (Result.FCentralDirectoryFile.CompressionMethod = ZipCompressionMethod_Deflate64) then
  begin
    case Result.FCentralDirectoryFile.GeneralPurposeBitFlag and 6 of
      0: Result.FCompressionType := ctNormal;
      2: Result.FCompressionType := ctMaximum;
      4: Result.FCompressionType := ctFast;
      6: Result.FCompressionType := ctSuperFast
    end;
  end;
  FParent.FIsDirty := True;
  if not FParent.FBatchMode then
  begin
    FParent.DoChange(FParent, 2);
  end;
end;

function TKAZipEntries.AddFolderChain(ItemName: string; FileAttr: Word;
  FileDate: TDateTime): Boolean;
var
  FN: string;
  TN: string;
  INCN: string;
  P: Integer;
  MS: TMemoryStream;
  NoMore: Boolean;
begin
  //  Result := False;
  FN := ExtractFilePath(ToDosName(ToZipName(ItemName)));
  TN := FN;
  INCN := '';
  MS := TMemoryStream.Create;
  try
    repeat
      NoMore := True;
      P := Pos('\', TN);
      if P > 0 then
      begin
        INCN := INCN + Copy(TN, 1, P);
        System.Delete(TN, 1, P);
        MS.Position := 0;
        MS.Size := 0;
        if IndexOf(INCN) = -1 then
        begin
          if FParent.FZipSaveMethod = FastSave then
            AddStreamFast(INCN, FileAttr, FileDate, MS)
          else if FParent.FZipSaveMethod = RebuildAll then
            AddStreamRebuild(INCN, FileAttr, FileDate, MS);
        end;
        NoMore := False;
      end;
    until NoMore;
    Result := True;
  finally
    MS.Free;
  end;
end;

function TKAZipEntries.AddFolderChain(ItemName: string): Boolean;
begin
  Result := AddFolderChain(ItemName, faDirectory, Now);
end;

function TKAZipEntries.AddStream(FileName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := nil;
  if (FParent.FStoreFolders) and (FParent.FStoreRelativePath) then
    AddFolderChain(FileName);

  if FParent.FZipSaveMethod = FastSave then
    Result := AddStreamFast(FileName, FileAttr, FileDate, Stream)
  else if FParent.FZipSaveMethod = RebuildAll then
    Result := AddStreamRebuild(FileName, FileAttr, FileDate, Stream);

  if Assigned(FParent.FOnAddItem) then
    FParent.FOnAddItem(FParent, FileName);
end;

function TKAZipEntries.AddStream(FileName: string; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := AddStream(FileName, faArchive, Now, Stream);
end;

function TKAZipEntries.AddFile(FileName, NewFileName: string): TKAZipEntriesEntry;
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

function TKAZipEntries.AddFile(FileName: string): TKAZipEntriesEntry;
begin
  Result := AddFile(FileName, FileName);
end;

function TKAZipEntries.AddFiles(FileNames: TStrings): Boolean;
var
  X: Integer;
begin
  Result := False;
  FParent.FBatchMode := True;
  try
    for X := 0 to FileNames.Count - 1 do
      AddFile(FileNames.Strings[X]);
  except
    FParent.FBatchMode := False;
    FParent.DoChange(FParent, 2);
    Exit;
  end;
  FParent.FBatchMode := False;
  FParent.DoChange(FParent, 2);
  Result := True;
end;

function TKAZipEntries.AddFolderEx(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
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
        if (FParent.FStoreFolders) and (FParent.FStoreRelativePath) then
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

function TKAZipEntries.AddFolder(FolderName: string; RootFolder: string; WildCard: string; WithSubFolders: Boolean): Boolean;
begin
  FParent.FBatchMode := True;
  try
    Result := AddFolderEx(FolderName, RootFolder, WildCard, WithSubFolders);
  finally
    FParent.FBatchMode := False;
    FParent.DoChange(FParent, 2);
  end;
end;

function TKAZipEntries.AddFilesAndFolders(FileNames: TStrings; RootFolder: string; WithSubFolders: Boolean): Boolean;
var
  X: Integer;
  Res: Integer;
  Dir: TSearchRec;
begin
  FParent.FBatchMode := True;
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
    FParent.FBatchMode := False;
    FParent.DoChange(FParent, 2);
  end;
  Result := True;
end;

procedure TKAZipEntries.RemoveFiles(List: TList);
begin
  if List.Count = 1 then
  begin
    Remove(Integer(List.Items[0]));
  end
  else
  begin
    SortList(List);
    FParent.FBatchMode := True;
    try
      RemoveBatch(List);
    finally
      FParent.FBatchMode := False;
      FParent.DoChange(Self, 3);
    end;
  end;
end;

procedure TKAZipEntries.RemoveSelected;
var
  X: Integer;
  List: TList;
begin
  FParent.FBatchMode := True;
  List := TList.Create;
  try
    for X := 0 to Count - 1 do
    begin
      if Self.Items[X].Selected then
        List.Add(Pointer(X));
    end;
    RemoveBatch(List);
  finally
    List.Free;
    FParent.FBatchMode := False;
    FParent.DoChange(Self, 3);
  end;
end;

procedure TKAZipEntries.ExtractToStream(Item: TKAZipEntriesEntry; TargetStream: TStream);
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
    FParent.FCurrentDFS := Item.SizeUncompressed;

    case Item.CompressionMethod of
      ZipCompressionMethod_Deflate:
        begin
          decompressor := TZDecompressionStream.Create(sourceStream);
          decompressor.OnProgress := FParent.OnDecompress;
          SetLength(BUF, FParent.FCurrentDFS);
          try
            NR := Decompressor.Read(BUF[1], FParent.FCurrentDFS);
            if NR = FParent.FCurrentDFS then
              targetStream.Write(BUF[1], FParent.FCurrentDFS);
          finally
            decompressor.Free;
          end;
        end;
{$IFDEF USE_BZIP2}
      ZipCompressionMethod_bzip2:
        begin
          DecompressorBZ2 := TBZDecompressionStream.Create(sourceStream);
          DecompressorBZ2.OnProgress := FParent.OnDecompress;
          SetLength(BUF, FParent.FCurrentDFS);
          try
            NR := DecompressorBZ2.Read(BUF[1], FParent.FCurrentDFS);
            if NR = FParent.FCurrentDFS then
              targetStream.Write(BUF[1], FParent.FCurrentDFS);
          finally
            DecompressorBZ2.Free;
          end;
        end
{$ENDIF}
      ZipCompressionMethod_Store:
        begin
          targetStream.CopyFrom(sourceStream, FParent.FCurrentDFS);
        end;
    end;
  finally
    sourceStream.Free;
  end;
end;

procedure TKAZipEntries.InternalExtractToFile(Item: TKAZipEntriesEntry; FileName: string);
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
    if FParent.FApplyAttributes then
    begin
      Attr := faArchive;
      if Item.FCentralDirectoryFile.ExternalFileAttributes and faHidden > 0 then
        Attr := Attr or faHidden;
      if Item.FCentralDirectoryFile.ExternalFileAttributes and faSysFile > 0 then
        Attr := Attr or faSysFile;
      if Item.FCentralDirectoryFile.ExternalFileAttributes and faReadOnly > 0 then
        Attr := Attr or faReadOnly;
      FileSetAttr(FileName, Attr);
    end;
  end;
end;

procedure TKAZipEntries.ExtractToFile(Item: TKAZipEntriesEntry; FileName: string);
var
  Can: Boolean;
  OA: TOverwriteAction;
begin
  OA := FParent.FOverwriteAction;
  Can := True;
  if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FParent.FOnOverwriteFile)) then
  begin
    if FileExists(FileName) then
    begin
      FParent.FOnOverwriteFile(FParent, FileName, OA);
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

procedure TKAZipEntries.ExtractToFile(ItemIndex: Integer; FileName: string);
var
  Can: Boolean;
  OA: TOverwriteAction;
begin
  OA := FParent.FOverwriteAction;
  Can := True;
  if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FParent.FOnOverwriteFile)) then
  begin
    if FileExists(FileName) then
    begin
      FParent.FOnOverwriteFile(FParent, FileName, OA);
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
    InternalExtractToFile(Items[ItemIndex], FileName);
end;

procedure TKAZipEntries.ExtractToFile(FileName, DestinationFileName: string);
var
  I: Integer;
  Can: Boolean;
  OA: TOverwriteAction;
begin
  OA := FParent.FOverwriteAction;
  Can := True;
  if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FParent.FOnOverwriteFile)) then
  begin
    if FileExists(DestinationFileName) then
    begin
      FParent.FOnOverwriteFile(FParent, DestinationFileName, OA);
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
    I := IndexOf(FileName);
    InternalExtractToFile(Items[I], DestinationFileName);
  end;
end;

procedure TKAZipEntries.ExtractAll(TargetDirectory: string);
var
  FN: string;
  DN: string;
  X: Integer;
  Can: Boolean;
  OA: TOverwriteAction;
  FileName: string;
begin
  OA := FParent.FOverwriteAction;
  Can := True;
  try
    for X := 0 to Count - 1 do
    begin
      FN := FParent.GetFileName(Items[X].FileName);
      DN := FParent.GetFilePath(Items[X].FileName);
      if DN <> '' then
        ForceDirectories(TargetDirectory + '\' + DN);
      FileName := TargetDirectory + '\' + DN + FN;
      if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FParent.FOnOverwriteFile)) then
      begin
        if FileExists(FileName) then
        begin
          FParent.FOnOverwriteFile(FParent, FileName, OA);
        end;
      end;
      case OA of
        oaSkip: Can := False;
        oaSkipAll: Can := False;
        oaOverwrite: Can := True;
        oaOverwriteAll: Can := True;
      end;
      if Can then
        InternalExtractToFile(Items[X], FileName);
    end;
  finally
  end;
end;

procedure TKAZipEntries.ExtractSelected(TargetDirectory: string);
var
  FN: string;
  DN: string;
  X: Integer;
  OA: TOverwriteAction;
  Can: Boolean;
  FileName: string;
begin
  OA := FParent.FOverwriteAction;
  Can := True;
  try
    for X := 0 to Count - 1 do
    begin
      if Items[X].FSelected then
      begin
        FN := FParent.GetFileName(Items[X].FileName);
        DN := FParent.GetFilePath(Items[X].FileName);
        if DN <> '' then
          ForceDirectories(TargetDirectory + '\' + DN);
        FileName := TargetDirectory + '\' + DN + FN;
        if ((OA <> oaOverwriteAll) and (OA <> oaSkipAll)) and (Assigned(FParent.FOnOverwriteFile)) then
        begin
          if FileExists(FileName) then
          begin
            FParent.FOnOverwriteFile(FParent, FileName, OA);
          end;
        end;
        case OA of
          oaSkip: Can := False;
          oaSkipAll: Can := False;
          oaOverwrite: Can := True;
          oaOverwriteAll: Can := True;
        end;
        if Can then
          InternalExtractToFile(Items[X], TargetDirectory + '\' + DN + FN);
      end;
    end;
  finally
  end;
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
  FParent.Rebuild;
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
var
  FN: string;
begin
  FN := IncludeTrailingBackslash(FolderName);
  AddFolderChain(FN, faDirectory, FolderDate);
  FParent.FIsDirty := True;
end;

procedure TKAZipEntries.RenameFolder(FolderName: string; NewFolderName: string);
var
  FN: string;
  NFN: string;
  S: string;
  X: Integer;
  L: Integer;
begin
  FN := ToZipName(IncludeTrailingBackslash(FolderName));
  NFN := ToZipName(IncludeTrailingBackslash(NewFolderName));
  L := Length(FN);
  if IndexOf(NFN) = -1 then
  begin
    for X := 0 to Count - 1 do
    begin
      S := Items[X].FileName;
      if Pos(FN, S) = 1 then
      begin
        System.Delete(S, 1, L);
        S := NFN + S;
        Items[X].FileName := S;
        FParent.FIsDirty := True;
      end;
    end;
    if (FParent.FIsDirty) and (FParent.FBatchMode = False) then
      Rebuild;
  end;
end;

procedure TKAZipEntries.RenameMultiple(Names: TStringList; NewNames: TStringList);
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
    FParent.FBatchMode := True;
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
      FParent.FBatchMode := False;
    end;
    if BR > 0 then
    begin
      Rebuild;
      FParent.DoChange(FParent, 6);
    end;
  end;
end;

function TKAZipEntries.AddEntryThroughStream(FileName: string; FileDate: TDateTime; FileAttr: Word): TStream;
var
  itemName: string;
  newEntry: TKAZIPEntriesEntry;
  i: Integer;
  zipComment: string;
  OBM: Boolean;
  newLocalEntryPosition: Int64;
begin
  if (FParent.FStoreFolders) and (FParent.FStoreRelativePath) then
    AddFolderChain(FileName);

  zipComment := FParent.Comment.Text;

  itemName := Filename;

  if not FParent.FStoreRelativePath then
    itemName := ExtractFileName(itemName);

  itemName := ToZipName(itemName);

  //If an item with this name already exists then remove it
  i := Self.IndexOf(itemName);
  if i >= 0 then
  begin
    OBM := FParent.FBatchMode;
    try
      if OBM = False then
        FParent.FBatchMode := True;
      Remove(i);
    finally
      FParent.FBatchMode := OBM;
    end;
  end;

  //This is where the new local entry starts (where the central directly is).
  //We overwrite the central direct (and EOCD marker) and then re-write them after the end of the file
  newLocalEntryPosition := FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory;

//  FParent.FCurrentDFS := uncompressedLength;

  //Fill records
  newEntry := TKAZipEntriesEntry(Self.Add);

  //Local file entry
  newEntry.FLocalFile.LocalFileHeaderSignature := SIG_LocalHeader; //'PK'#03#04
  newEntry.FLocalFile.VersionNeededToExtract := 20;
  newEntry.FLocalFile.GeneralPurposeBitFlag := 0;
  newEntry.FLocalFile.CompressionMethod := ZipCompressionMethod_Deflate; //assume Deflate, but if the length is zero then it will be switched to Store (0)
  newEntry.FLocalFile.LastModFileTimeDate := DateTimeToFileDate(FileDate);
  newEntry.FLocalFile.Crc32 := 0; //don't know it yet, will back-fill
  newEntry.FLocalFile.CompressedSize := 0; //don't know it yet, will back-fill
  newEntry.FLocalFile.UncompressedSize := 0; //don't know it yet, will back-fill
  newEntry.FLocalFile.FilenameLength := Length(ItemName);
  newEntry.FLocalFile.ExtraFieldLength := 0;
  newEntry.FLocalFile.FileName := ItemName;
  newEntry.FLocalFile.ExtraField := '';
  newEntry.FLocalFile.CompressedData := ''; //not used

  //Create the Central Directory entry
  newEntry.FCentralDirectoryFile.CentralFileHeaderSignature := SIG_CentralFile; //PK 0x01 0x02
  newEntry.FCentralDirectoryFile.VersionMadeBy := 20;
  newEntry.FCentralDirectoryFile.VersionNeededToExtract := 20;
  newEntry.FCentralDirectoryFile.GeneralPurposeBitFlag := 0;
  newEntry.FCentralDirectoryFile.CompressionMethod := ZipCompressionMethod_Deflate; //assume Deflate, but if the length is zero then it will be switched to Store (0)
  newEntry.FCentralDirectoryFile.LastModFileTimeDate := DateTimeToFileDate(FileDate);
  newEntry.FCentralDirectoryFile.Crc32 := 0; //don't know it yet, will back-fill
  newEntry.FCentralDirectoryFile.CompressedSize := 0; //don't know it yet, will back-fill
  newEntry.FCentralDirectoryFile.UncompressedSize := 0; //don't know it yet, will back-fill
  newEntry.FCentralDirectoryFile.FilenameLength := Length(ItemName);
  newEntry.FCentralDirectoryFile.ExtraFieldLength := 0;
  newEntry.FCentralDirectoryFile.FileCommentLength := 0;
  newEntry.FCentralDirectoryFile.DiskNumberStart := 0;
  newEntry.FCentralDirectoryFile.InternalFileAttributes := 0;
  newEntry.FCentralDirectoryFile.ExternalFileAttributes := FileAttr;
  newEntry.FCentralDirectoryFile.RelativeOffsetOfLocalHeader := newLocalEntryPosition;
  newEntry.FCentralDirectoryFile.FileName := ItemName;
  newEntry.FCentralDirectoryFile.ExtraField := '';
  newEntry.FCentralDirectoryFile.FileComment := '';

  newEntry.FDate := FileDate;

  if (newEntry.FCentralDirectoryFile.GeneralPurposeBitFlag and 1) > 0 then
    newEntry.FIsEncrypted := True
  else
    newEntry.FIsEncrypted := False;
  newEntry.FIsFolder := (newEntry.FCentralDirectoryFile.ExternalFileAttributes and faDirectory) > 0;


  // Write Local file entry to stream
  FParent.FZipStream.Position := newLocalEntryPosition;
  FParent.FZipStream.Write(newEntry.FLocalFile, SizeOf(newEntry.FLocalFile) - 3 * SizeOf(string)); //excluding FileName, ExtraField, and CompressedData
  if newEntry.FLocalFile.FilenameLength > 0 then
    FParent.FZipStream.Write(newEntry.FLocalFile.FileName[1], newEntry.FLocalFile.FilenameLength);

  Result := TKAZipStream.Create(FParent.FZipStream, newEntry, FParent.FZipCompressionType, FParent);
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
  FZipHeader.ParseZip(MS);
  FIsZipFile := FZipHeader.FIsZipFile;
  if not FIsZipFile then
    Close;
  FIsDirty := True;
  DoChange(Self, 1);
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
  DoChange(Self, 0);
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
begin
  Entries.ExtractToFile(Item, FileName);
end;

procedure TKAZip.ExtractToFile(ItemIndex: Integer; FileName: string);
begin
  Entries.ExtractToFile(ItemIndex, FileName);
end;

procedure TKAZip.ExtractToFile(FileName, DestinationFileName: string);
begin
  Entries.ExtractToFile(FileName, DestinationFileName);
end;

procedure TKAZip.ExtractToStream(Item: TKAZipEntriesEntry; Stream: TStream);
begin
  Entries.ExtractToStream(Item, Stream);
end;

procedure TKAZip.ExtractAll(TargetDirectory: string);
begin
  Entries.ExtractAll(TargetDirectory);
end;

procedure TKAZip.ExtractSelected(TargetDirectory: string);
begin
  Entries.ExtractSelected(TargetDirectory);
end;

function TKAZip.AddFile(FileName, NewFileName: string): TKAZipEntriesEntry;
begin
  Result := Entries.AddFile(FileName, NewFileName);
end;

function TKAZip.AddFile(FileName: string): TKAZipEntriesEntry;
begin
  Result := Entries.AddFile(FileName);
end;

function TKAZip.AddFiles(FileNames: TStrings): Boolean;
begin
  Result := Entries.AddFiles(FileNames);
end;

function TKAZip.AddFolder(FolderName, RootFolder, WildCard: string;
  WithSubFolders: Boolean): Boolean;
begin
  Result := Entries.AddFolder(FolderName, RootFolder, WildCard, WithSubFolders);
end;

function TKAZip.AddFilesAndFolders(FileNames: TStrings; RootFolder: string;
  WithSubFolders: Boolean): Boolean;
begin
  Result := Entries.AddFilesAndFolders(FileNames, RootFolder, WithSubFolders);
end;

function TKAZip.AddStream(FileName: string; FileAttr: Word; FileDate: TDateTime; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := Entries.AddStream(FileName, FileAttr, FileDate, Stream);
end;

function TKAZip.AddStream(FileName: string; Stream: TStream): TKAZipEntriesEntry;
begin
  Result := Entries.AddStream(FileName, Stream);
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
  Entries.RemoveFiles(List);
end;

procedure TKAZip.RemoveSelected;
begin
  Entries.RemoveSelected;
  ;
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
  LF: TLocalFile;
begin
  //************************************************* RESAVE ALL LOCAL BLOCKS
  SetLength(NewLHOffsets, Entries.Count + 1);
  for X := 0 to Entries.Count - 1 do
  begin
    NewLHOffsets[X] := MS.Position;
    LF := Entries.GetLocalEntry(FZipStream, Entries.Items[X].LocalOffset, False);
    MS.Write(LF, SizeOf(LF) - 3 * SizeOf(string));
    if LF.FilenameLength > 0 then
      MS.Write(LF.FileName[1], LF.FilenameLength);
    if LF.ExtraFieldLength > 0 then
      MS.Write(LF.ExtraField[1], LF.ExtraFieldLength);
    if LF.CompressedSize > 0 then
      MS.Write(LF.CompressedData[1], LF.CompressedSize);
    if Assigned(FOnRebuildZip) then
      FOnRebuildZip(Self, X, Entries.Count - 1);
  end;
end;

procedure TKAZip.RebuildCentralDirectory(MS: TStream);
var
  X: Integer;
  CDF: TCentralDirectoryFile;
begin
  NewEndOfCentralDir := FEndOfCentralDir;
  NewEndOfCentralDir.TotalNumberOfEntriesOnThisDisk := Entries.Count;
  NewEndOfCentralDir.TotalNumberOfEntries := Entries.Count;
  NewEndOfCentralDir.OffsetOfStartOfCentralDirectory := MS.Position;
  for X := 0 to Entries.Count - 1 do
  begin
    CDF := Entries.Items[X].FCentralDirectoryFile;
    CDF.RelativeOffsetOfLocalHeader := NewLHOffsets[X];
    MS.Write(CDF, SizeOf(CDF) - 3 * SizeOf(string));
    if CDF.FilenameLength > 0 then
      MS.Write(CDF.FileName[1], CDF.FilenameLength);
    if CDF.ExtraFieldLength > 0 then
      MS.Write(CDF.ExtraField[1], CDF.ExtraFieldLength);
    if CDF.FileCommentLength > 0 then
      MS.Write(CDF.FileComment[1], CDF.FileCommentLength);
    if Assigned(FOnRebuildZip) then
      FOnRebuildZip(Self, X, Entries.Count - 1);
  end;
  NewEndOfCentralDir.SizeOfTheCentralDirectory := Cardinal(MS.Position) - NewEndOfCentralDir.OffsetOfStartOfCentralDirectory;
end;

procedure TKAZip.RebuildEndOfCentralDirectory(MS: TStream);
var
  ZipComment: string;
begin
  ZipComment := Comment.Text;
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

procedure TKAZip.FixZip(MS: TStream);
var
  X: Integer;
  Y: Integer;
  NewCount: Integer;
  LF: TLocalFile;
  CDF: TCentralDirectoryFile;
  ZipComment: string;
begin
  ZipComment := Comment.Text;
  Y := 0;
  SetLength(NewLHOffsets, Entries.Count + 1);
  for X := 0 to Entries.Count - 1 do
  begin
    LF := Entries.GetLocalEntry(FZipStream, Entries.Items[X].LocalOffset, False);
    if (LF.LocalFileHeaderSignature = $04034B50) and (Entries.Items[X].Test) then
    begin
      NewLHOffsets[Y] := MS.Position;
      MS.Write(LF, SizeOf(LF) - 3 * SizeOf(string));
      if LF.FilenameLength > 0 then
        MS.Write(LF.FileName[1], LF.FilenameLength);
      if LF.ExtraFieldLength > 0 then
        MS.Write(LF.ExtraField[1], LF.ExtraFieldLength);
      if LF.CompressedSize > 0 then
        MS.Write(LF.CompressedData[1], LF.CompressedSize);
      if Assigned(FOnRebuildZip) then
        FOnRebuildZip(Self, X, Entries.Count - 1);
      Inc(Y);
    end
    else
    begin
      Entries.Items[X].FCentralDirectoryFile.CentralFileHeaderSignature := 0;
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
    CDF := Entries.Items[X].FCentralDirectoryFile;
    if CDF.CentralFileHeaderSignature = SIG_CentralFile then
    begin
      CDF.RelativeOffsetOfLocalHeader := NewLHOffsets[Y];
      MS.Write(CDF, SizeOf(CDF) - 3 * SizeOf(string));
      if CDF.FilenameLength > 0 then
        MS.Write(CDF.FileName[1], CDF.FilenameLength);
      if CDF.ExtraFieldLength > 0 then
        MS.Write(CDF.ExtraField[1], CDF.ExtraFieldLength);
      if CDF.FileCommentLength > 0 then
        MS.Write(CDF.FileComment[1], CDF.FileCommentLength);
      if Assigned(FOnRebuildZip) then
        FOnRebuildZip(Self, X, Entries.Count - 1);
      Inc(Y);
    end;
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
  RebuildEndOfCentralDirectory(Stream);
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
begin
  Entries.RenameMultiple(Names, NewNames);
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
begin
  Entries.CreateFolder(FolderName, FolderDate);
end;

procedure TKAZip.RenameFolder(FolderName: string; NewFolderName: string);
begin
  Entries.RenameFolder(FolderName, NewFolderName);
end;

procedure TKAZip.SetReadOnly(const Value: Boolean);
begin
  FReadOnly := Value;
end;

procedure TKAZip.SetApplyAtributes(const Value: Boolean);
begin
  FApplyAttributes := Value;
end;

procedure TKAZip.WriteCentralDirectory(TargetStream: TStream);
var
  i: Integer;
  entry: TKAZipEntriesEntry;
begin
  //Save the Central Directory entries
  for i := 0 to Self.Entries.Count-1 do
  begin
    entry := Self.Entries.Items[i];

    //packed byte data
    FZipStream.Write(entry.FCentralDirectoryFile, SizeOf(entry.FCentralDirectoryFile) - 3 * SizeOf(string));

    //optional filenmae
    if entry.FCentralDirectoryFile.FilenameLength > 0 then
      FZipStream.Write(entry.FCentralDirectoryFile.FileName[1], entry.FCentralDirectoryFile.FilenameLength);

    //optional extra field
    if entry.FCentralDirectoryFile.ExtraFieldLength > 0 then
      FZipStream.Write(entry.FCentralDirectoryFile.ExtraField[1], entry.FCentralDirectoryFile.ExtraFieldLength);

    //optional file comment
    if entry.FCentralDirectoryFile.FileCommentLength > 0 then
      FZipStream.Write(entry.FCentralDirectoryFile.FileComment[1], entry.FCentralDirectoryFile.FileCommentLength);
  end;
end;

function TKAZip.AddEntryThroughStream(FileName: string; FileDate: TDateTime; FileAttr: Word): TStream;
begin
  Result := Entries.AddEntryThroughStream(FileName, FileDate, FileAttr);
end;

function TKAZip.AddEntryThroughStream(FileName: string): TStream;
begin
  Result := Entries.AddEntryThroughStream(FileName, Now, faArchive);
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
  data: PByteArray;
begin
  state := FCRC32;

  data := Addr(Buffer);

  for i := 0 to Count-1 do
    state := (state shr 8) xor CRCTable[Byte(state) xor data[i]];

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
  zipComment: string;
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
  FEntry.FLocalFile.Crc32 := crc32;
  FEntry.FLocalFile.UncompressedSize := uncompressedLength;
  FEntry.FLocalFile.CompressedSize := compressedLength;
  FEntry.FLocalFile.CompressionMethod := compressionMethod;
  FEntry.FCentralDirectoryFile.Crc32 := crc32;
  FEntry.FCentralDirectoryFile.UncompressedSize := uncompressedLength;
  FEntry.FCentralDirectoryFile.CompressedSize := compressedLength;
  FEntry.FCentralDirectoryFile.CompressionMethod := compressionMethod;

  //Re-write the (now complete) local header
  try
    FTargetStream.Position := FEntry.FCentralDirectoryFile.RelativeOffsetOfLocalHeader;
    FTargetStream.Write(FEntry.FLocalFile, SizeOf(FEntry.FLocalFile) - 3 * SizeOf(string)); //excluding FileName, ExtraField, and CompressedData
  finally
    //Return to where the central directory will be, and write it
    FTargetStream.Position := centralDirectoryPosition;
  end;

  //Save the Central Directory entries (todo: split this into separate function)
  FParent.WriteCentralDirectory(FParent.FZipStream);

  //Save EOCD (End of Central Directory) record
  FParent.FEndOfCentralDirPos := FParent.FZipStream.Position;
  FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory := centralDirectoryPosition;
  FParent.FEndOfCentralDir.SizeOfTheCentralDirectory := FParent.FEndOfCentralDirPos - FParent.FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
  Inc(FParent.FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
  Inc(FParent.FEndOfCentralDir.TotalNumberOfEntries);
  FParent.FZipStream.Write(FParent.FEndOfCentralDir, SizeOf(TEndOfCentralDir));

  //Save ZIP comment (if any)
  zipComment := FParent.Comment.Text;
  FParent.FZipCommentPos := FParent.FZipStream.Position;
  if Length(zipComment) > 0 then
    FParent.FZipStream.Write(zipComment[1], Length(zipComment));


  if (FEntry.FCentralDirectoryFile.CompressionMethod = ZipCompressionMethod_Deflate) or (FEntry.FCentralDirectoryFile.CompressionMethod = ZipCompressionMethod_Deflate64) then
  begin
    case FEntry.FCentralDirectoryFile.GeneralPurposeBitFlag and 6 of
      0: FEntry.FCompressionType := ctNormal;
      2: FEntry.FCompressionType := ctMaximum;
      4: FEntry.FCompressionType := ctFast;
      6: FEntry.FCompressionType := ctSuperFast
      else
        FEntry.FCompressionType := ctUnknown;
    end;
  end
  else
    FEntry.FCompressionType := ctUnknown;



  FParent.FIsDirty := True;
  if not FParent.FBatchMode then
  begin
    FParent.DoChange(FParent, 2);
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

