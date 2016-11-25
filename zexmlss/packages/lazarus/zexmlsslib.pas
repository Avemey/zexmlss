{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit zexmlsslib;

interface

uses
  zexmlss, zexmlssutils, zsspxml, zeformula, zeodfs, zesavecommon, zeSave, 
  zeSaveEXML, zeSaveODS, zeSaveXLSX, zeZippy, zeZippyLazip, zexlsx, 
  zearchhelper, zenumberformats, LazarusPackageIntf;

implementation

procedure Register;
begin
  RegisterUnit('zexmlss', @zexmlss.Register);
end;

initialization
  RegisterPackage('zexmlsslib', @Register);
end.
