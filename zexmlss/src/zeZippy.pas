unit zeZippy;
(* Simplistic interface to creating simplistic Zip files.
   Bridge object template for avemey.com components.

   (c) the Arioch, licensed under zLib license *)

interface uses Classes, SysUtils;

type
///  Life span: create, add files, save to disk, free
///  State reflects them.
   TZxZipGenState = (zgsAccepting, zgsFlushing, zgsSealed);

   EZxZipGen = class(Exception);

   TZxZipGen = class;
   CZxZipGens = class of TZxZipGen;
   TZxZipGen = class  // abstract   D7 does not support the word, and it actually does not mean much
     public
       /// creates new empty generator, ready to accept files
       /// Implementations should try to create the file for writing
       ///    and throw exception if they can not.
       constructor Create(const ZipFile: TFileName); virtual;

       /// gives new stream for the internal file
       /// Exporter should stuff it with data and then seal it.
       /// Exporter should not assume ownership of the stream.
       function NewStream(const RelativeName: TFileName): TStream;

       /// Communicates the generator, that internal file is complete.
       /// Depending on implementation, that can flush the file and free
       ///   the stream object, or keep it in memory and wait for later.
       /// Anyway, after the call this stream should not be used, it may
       ///   already be freed.
       /// Only streams given by NewStream are accepted.
       procedure SealStream(const Data: TStream);

       procedure AbortAndDelete;
       procedure SaveAndSeal;

     protected
       FActiveStreams, FSealedStreams: TStringList;

       procedure DoAbortAndDelete;  virtual;
       procedure DoSaveAndSeal;  virtual; abstract;
       function  DoNewStream(const RelativeName: TFileName): TStream; virtual;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelativeName: TFileName): boolean;  virtual; abstract;

     protected
       FFileName: TFileName;
       FState: TZxZipGenState;
       procedure RequireState(const st: TZxZipGenState);
       procedure ChangeState(const _From, _To: TZxZipGenState);

     public
       property State: TZxZipGenState read FState;
       property ZipFileName: TFileName read FFileName;

     public
       procedure AfterConstruction; override;
       procedure BeforeDestruction; override;

     protected // private ?  strict private ?
        class procedure RegisterZipGen(const zgc: CZxZipGens);
        class procedure UnRegisterZipGen(const zgc: CZxZipGens);

     public
        class procedure Register;
        class procedure UnRegister;

        /// The spreadsheet exporter may enumerate them increasing parameter until nil is returned
        class function QueryZipGen(const idx: integer = 0): CZxZipGens;

        /// Returns the empty "fall-back" class that just makes a folder with uncompressed files
        /// You may register it if you want ;-)
        class function QueryDummyZipGen: CZxZipGens;
   end;

implementation uses TypInfo, Contnrs;

{ TZxZipGen }

procedure TZxZipGen.AfterConstruction;
  procedure Tune(const sl: TStringList);
  begin
    sl.CaseSensitive := false;
    sl.Duplicates    := dupError;
    sl.Sorted        := true;
  end;
begin
  inherited;

  FSealedStreams := TStringList.Create;
  FActiveStreams := TStringList.Create;
  Tune(FSealedStreams);
  Tune(FActiveStreams);
end;

procedure TZxZipGen.BeforeDestruction;
  procedure Wipe(const sl: TStringList); var i: integer;
  begin
    for i := sl.Count - 1 downto 0 do begin
        sl.Objects[i].Free; // can be nil - that is okay for .Free
       // sl.Delete(i); like anyone cares, really :-)
    end;
    sl.Free;
  end;

begin
  Wipe(FSealedStreams);
  Wipe(FActiveStreams);

  inherited;

  if State <> zgsSealed then begin
     FState := zgsSealed; // breaking infinite loop destructor -> exception -> destructor -> exception ->....
     raise EZxZipGen.Create('Zip file should be either saved or destroyed before freeing');
  end;
// even if no streams were added - the file garbage may remain
end;

procedure TZxZipGen.AbortAndDelete;
begin
  ChangeState(zgsAccepting, zgsFlushing);

  try
    DoAbortAndDelete;
  finally
    ChangeState(zgsFlushing, zgsSealed);
  end;
end;

procedure TZxZipGen.SaveAndSeal;
begin
  RequireState(zgsAccepting);
  if FActiveStreams.Count > 0 then
     raise EZxZipGen.Create('All data streams should be sealed before saving Zip to disk.');

  ChangeState(zgsAccepting, zgsFlushing);
  try
    DoSaveAndSeal;
  finally
    ChangeState(zgsFlushing, zgsSealed);
  end;
end;

procedure TZxZipGen.SealStream(const Data: TStream);
var idx: Integer;
begin
  RequireState(zgsAccepting);

  if nil = Data then
     raise EZxZipGen.Create('Given data is nil.');

  idx := FActiveStreams.IndexOfObject(Data);
  if idx < 0 then
     raise EZxZipGen.Create('Given data (' + Data.ClassName + ') does not belong to this zip generator.');

  if DoSealStream(Data, FActiveStreams[idx]) then begin
     Data.Free;
     FSealedStreams.Add(FActiveStreams.Strings[idx]); // should remember to not add it twice
     FActiveStreams.Delete(idx);
  end else begin
     FSealedStreams.AddObject(FActiveStreams.Strings[idx], FActiveStreams.Objects[idx]);
     FActiveStreams.Delete(idx);
  end;
end;

function TZxZipGen.NewStream(const RelativeName: TFileName): TStream;
var idx: integer;
    fname: TFileName;
 procedure UnifyDelims(const c: char; const PathDelim: char = SysUtils.PathDelim);
 begin
  if PathDelim <> c then
     fname := StringReplace(fname, c, PathDelim, [rfReplaceAll]);

  while Pos(PathDelim + c, fname)>0 do // ugly, but actually should bever happen! unless bugcheck
        fname := StringReplace(fname, PathDelim+PathDelim, PathDelim, [rfReplaceAll]);
 end;
begin
  RequireState(zgsAccepting);

  if RelativeName[Length(RelativeName)] = PathDelim then
     raise EZxZipGen.Create('Zip generator accepts files, not directories: '+RelativeName);

  fname := RelativeName;
//  UnifyDelims('\'); UnifyDelims('/');
  UnifyDelims('\', '/');  // Excel 2010 chokes on back-slashes

  if fname = '' then raise EZxZipGen.Create('Internal file should have a name.');

  if fname[1] = '/' then Delete(fname,1,1);

  if FSealedStreams.Find(fname, idx) or FActiveStreams.Find(fname, idx) then
     raise EZxZipGen.Create('File >>'+fname+'<< already was added to generator.');

  Result := DoNewStream(fname);
  FActiveStreams.AddObject(fname, Result);
end;

procedure TZxZipGen.RequireState(const st: TZxZipGenState);
  function Name(const st: TZxZipGenState): string;
  begin
    Result := GetEnumName( TypeInfo(TZxZipGenState), ord(st) );
  end;
begin
  if State <> st then
     raise EZxZipGen.CreateFmt(
       'Zip generator state is %s while %s is required for further processing.',
       [ Name(State), Name(st) ]);
end;

procedure TZxZipGen.ChangeState(const _From, _To: TZxZipGenState);
begin
  RequireState(_From);
  FState := _To;
end;

constructor TZxZipGen.Create(const ZipFile: TFileName);
begin
   inherited Create;
   FFileName := ZipFile;
end;

procedure TZxZipGen.DoAbortAndDelete; // var i: integer;
begin
//  for i := FActiveStreams.Count - 1 downto 0 do begin
//      FActiveStreams.Objects[i].Free; // can be nil - that is okay for .Free
//      FActiveStreams.Delete(i);
//  end;
  FActiveStreams.OwnsObjects := True; // force freeing them
  FActiveStreams.Clear;
end;

function TZxZipGen.DoNewStream(const RelativeName: TFileName): TStream;
begin
  Result := TMemoryStream.Create;
end;

(*********************************************************)

type TZxFolderInsteadOfZip = class (TZxZipGen)
     public

       /// Implementations should try to create the file for writing
       ///    and throw exception if they can not.
       constructor Create(const ZipFile: TFileName); override;

     protected
       function MakeAbsPath(const RelativeName: TFileName): TFileName;

       procedure DoAbortAndDelete;  override;
       procedure DoSaveAndSeal;  override;
       function  DoNewStream(const RelativeName: TFileName): TStream; override;

       /// Returns True if the stream was flushed and clearance is given to free it.
       /// Otherwise is transmitted to the sealed list
       function  DoSealStream(const Data: TStream; const RelName: TFileName): boolean;  override;
   end;

{ TZxFolderInsteadOfZip }

constructor TZxFolderInsteadOfZip.Create(const ZipFile: TFileName);
begin
  inherited;

  if not ForceDirectories(ZipFileName) // better use name supplied by the base class
     then raise EZxZipGen.Create('Cannot create folder '+ZipFileName);
end;

procedure TZxFolderInsteadOfZip.DoSaveAndSeal;
begin
  // nothing to do - everythign was done on the go
end;

procedure TZxFolderInsteadOfZip.DoAbortAndDelete;
var s: string; i: integer;
begin
  for i := FActiveStreams.Count - 1 downto 0 do begin
      FActiveStreams.Objects[i].Free; // can be nil - that is okay for .Free

      s := FActiveStreams.Strings[i];
      DeleteFile( MakeAbsPath(s));

      repeat
        s := ExtractFileDir(s);
// "This string is empty if FileName contains no drive and directory parts."

        if '' = s  then break;
        if not RemoveDir(s) then break; // error. Parents would not be releted too.
      until '' = s; // redundancy but safety

      FActiveStreams.Delete(i);
  end;
  RmDir(ZipFileName);
end;

function TZxFolderInsteadOfZip.MakeAbsPath(const RelativeName: TFileName):TFileName;
begin
  Result := ZipFileName + PathDelim + RelativeName
end;

function TZxFolderInsteadOfZip.DoNewStream(
  const RelativeName: TFileName): TStream;
var fn: TFileName;
begin
  fn := MakeAbsPath(RelativeName);
  ForceDirectories(ExtractFileDir( fn ));

  Result := TFileStream.Create(fn, fmCreate, fmShareExclusive );
end;

function TZxFolderInsteadOfZip.DoSealStream(const Data: TStream; const RelName: TFileName): boolean;
begin
  Result := True; // data is saved on the go - just free it
end;

(*********************************************************)

var ZxZipMakers: TClassList;

class procedure TZxZipGen.Register;
begin
  RegisterZipGen(Self);
end;

class procedure TZxZipGen.RegisterZipGen(const zgc: CZxZipGens);
begin
  if (nil = zgc) or not zgc.InheritsFrom(TZxZipGen) then
     raise EZxZipGen.Create('Can only register zip generating subclasses.');

  ZxZipMakers.Insert(0, zgc); // LIFO
end;

class procedure TZxZipGen.UnRegister;
begin
  UnRegisterZipGen(Self);
end;

class procedure TZxZipGen.UnRegisterZipGen(const zgc: CZxZipGens);
begin
  ZxZipMakers.Remove(zgc);
end;

/// The spreadsheet exporter may enumerate them increasing parameter until nil is returned
class function TZxZipGen.QueryZipGen(const idx: integer): CZxZipGens;
var c: TClass;
begin
  Result := nil;
  if idx < 0 then exit;
  if idx >= ZxZipMakers.Count then exit;

  c := ZxZipMakers[idx];
  if not c.InheritsFrom(TZxZipGen) then
     raise EZxZipGen.Create(c.ClassName + ' was registered instead of Zip generator');

  Result := CZxZipGens(c);
end;

/// Returns the empty "fall-back" class that just makes a folder with uncompressed files
/// You may register it if you want ;-)
class function TZxZipGen.QueryDummyZipGen: CZxZipGens;
begin
  Result := TZxFolderInsteadOfZip;
end;



initialization
  ZxZipMakers := TClassList.Create;
finalization
  ZxZipMakers.Free;
end.
