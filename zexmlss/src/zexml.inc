
// ≈сли не используетс€ ZColorStringGrid 
{$DEFINE NOZCOLORSTRINGGRID}

// Free Pascal / Lazarus
{$IFDEF FPC}
  //дл€ FPC пока нету ZColorStringGrid-а
  {$DEFINE NOZCOLORSTRINGGRID}
  //≈сли хотим использовать модуль перекодировки LConvEncoding.pas
  {$DEFINE USELCONVENCODING}
{$ELSE}
  {$UNDEF USELCONVENCODING}
{$ENDIF}