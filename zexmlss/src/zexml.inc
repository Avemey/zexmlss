// If not using ZColorStringGrid
// Если не используется ZColorStringGrid 
{$DEFINE NOZCOLORSTRINGGRID}

// Conditional formatting
// Условное форматирование
{$DEFINE ZUSE_CONDITIONAL_FORMATTING}

// Компактная запись XML, без отступов и переводов строк
// Compact XML without indents and newlines
{$DEFINE NOXMLINDENTS}

// Charts
// Графики
{$DEFINE ZUSE_CHARTS}

// Drawings
// Рисунки
{$DEFINE ZUSE_DRAWINGS}

// Archivers for Delphi, not used for Free Pascal / Lazarus
// Архиваторы для Delphi, не используются для Free Pascal / Lazarus
{ use Delphi XE2 and above }
{$if CompilerVersion >= 23.0}         // bds xe2 (2012)
  {$define XE2ZIP}
{$else}
  { use KaZip (recommended) ) }
  {$define KAZIP}

  { use JEDI Code Library (via 7z) }
  {//$define JCL7Z}

  { use TurboPower Abbrevia }
  {//$define ABZIP}

  { use SciZipFile }
  {//$define SCIZIP}

  { use Synopse }
  {//$define SYNZIP}
{$ifend}

// Free Pascal / Lazarus
{$IFDEF FPC}
  //для FPC пока нету ZColorStringGrid-а
  {$DEFINE NOZCOLORSTRINGGRID}
  //Если хотим использовать модуль перекодировки LConvEncoding.pas
  {$DEFINE USELCONVENCODING}
{$ELSE}
  {$UNDEF USELCONVENCODING}
{$ENDIF}