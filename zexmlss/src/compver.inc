{$DEFINE UNKNOWN_COMPILER}
{$WARN IMPLICIT_STRING_CAST OFF}                    // Only care on Loss

{$IFDEF FPC} //KO: FPC
  {$UNDEF UNKNOWN_COMPILER}
  //Directive deprecated with string msg (deprecated 'Use "some else"';)
  //  works in FPC and D>=2009.
  {$DEFINE USE_DEPRECATED_STRING}
  {$DEFINE FPC_OR_DELPHI_UNICODE}
  {$I fpccompver.inc}
{$ENDIF}

/// ******DELPHI*********

{$IFDEF VER80} //delphi 1
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER90} //delphi 2
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER100} //delphi 3
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER120} //delphi 4
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER130} //delphi 5
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER140} //delphi 6
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER150} //delphi 7
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER160} //delphi 8
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER170} //delphi 9 (2005)
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER180} //delphi 10 (BDS2006)
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER185} //delphi 11 (2007)
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER190} // Delphi 2007 for .NET
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

/// ******Cbuilder*********

{$IFDEF VER93} //cbuilder1
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER110} //cbuilder 3
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF VER125} //cbuilder 4
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

/// ****** Delphi unicode *********

{$IFDEF VER200} // RAD Studio 2009
  {$UNDEF UNKNOWN_COMPILER}
  {$DEFINE DELPHI_UNICODE}
{$ENDIF VER200}

{$IFDEF VER210} // RAD Studio 2010
  {$UNDEF UNKNOWN_COMPILER}
  {$DEFINE DELPHI_UNICODE}
{$ENDIF VER210}

{$IFDEF VER220} // RAD Studio XE
  {$UNDEF UNKNOWN_COMPILER}
  {$DEFINE DELPHI_UNICODE}
{$ENDIF VER220}

{$IFDEF VER230} // RAD Studio XE2 
  {$UNDEF UNKNOWN_COMPILER}
  {$DEFINE DELPHI_UNICODE}
{$ENDIF VER230}

{$IFDEF UNKNOWN_COMPILER} // > XE2 
  {$DEFINE DELPHI_UNICODE}
  {$UNDEF UNKNOWN_COMPILER}
{$ENDIF}

{$IFDEF DELPHI_UNICODE}
  {$DEFINE USE_DEPRECATED_STRING}
  {$DEFINE FPC_OR_DELPHI_UNICODE}
  {$DEFINE Z_USE_FORMAT_SETTINGS}  //Have Formatsettings
{$ENDIF}

{$IFDEF VER200} // RAD Studio 2009
  {$UNDEF Z_USE_FORMAT_SETTINGS} //There are no FormatSettings in D2009!!!!
{$ENDIF VER200}

{ Embarcadero Delphi 24 XE10.1 (Berlin) }
{$IFDEF VER310} {$DEFINE XE10_1} {$DEFINE BERLIN}{$ENDIF}
{$IF CompilerVersion >= 31}   // Define the Plus (XP_UP) and Pre (PRE_XP) versions
  {$DEFINE XE10_1_UP} {$DEFINE BERLIN_UP}
  {$DEFINE TARRAY_HELPERFIXED}  // starting in Berlin these are handled correctly, prior to that all TArray<T> used a single Helper for all Ts
{$ELSE}
  {$DEFINE PRE_XE10_1} {$DEFINE PRE_BERLIN}
{$IFEND}

{ $DEFINE DELPHI_UNICODE}
