{ This file was automatically created by Lazarus. Do not edit!
  This source is only used to compile and install the package.
 }

unit zclabel; 

interface

uses
  zclab, zcftext, LazarusPackageIntf;

implementation

procedure Register; 
begin
  RegisterUnit('zclab', @zclab.Register); 
end; 

initialization
  RegisterPackage('zclabel', @Register); 
end.
