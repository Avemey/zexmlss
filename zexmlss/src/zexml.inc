
// Если не используется ZColorStringGrid 
{ $DEFINE NOZCOLORSTRINGGRID}

// Условное форматирование
{$DEFINE ZUSE_CONDITIONAL_FORMATTING}

//Charts
{$DEFINE ZUSE_CHARTS}

// Free Pascal / Lazarus
{$IFDEF FPC}
  //для FPC пока нету ZColorStringGrid-а
  {$DEFINE NOZCOLORSTRINGGRID}
  //Если хотим использовать модуль перекодировки LConvEncoding.pas
  {$DEFINE USELCONVENCODING}
{$ELSE}
  {$UNDEF USELCONVENCODING}
{$ENDIF}