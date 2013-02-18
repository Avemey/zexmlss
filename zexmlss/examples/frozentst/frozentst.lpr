program frozentst;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Classes, SysUtils, frozproc
  { you can add units after this },
  //lclproc, translations;
  interfaces;

begin
  TestSplit();
end.

