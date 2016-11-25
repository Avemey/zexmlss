ZEXMLSSLIB 0.0.15 (beta)
Для Lazarus, Delphi 7, C++Builder 6.
Borland Developer Studio 2005, BDS 2006, CodeGear Delphi 2007, 
CodeGear RAD Studio 2009, 2010, Delphi XE і Delphi XE2.
(Delphi 5 і 6 не правяраў. Delphi 5 толькі з кадоўкамі windows-1251 і CP866, але можна дадаць
свае функцыі)


Праца з файламі Excel XML Spreadsheet / OpenDocument Format (ODS) / OOXML (xlsx) без усталяванага Excel-а/OO Calc-а.

Усталёўка
---------
 1. Lazarus
  0. Калі не выкарыстоўваеце модуль перакадоўкі LConvEncoding.pas, прыбярыце 
     ў src\zexml.inc радок {$DEFINE USELCONVENCODING} (з UTF-8 будзе 
     працаваць звычайна, а астатнія кадоўкі прыйдзецца ручкамі)
  1. Пакет->Адкрыць файл	пакета (*.lpk) (Components->Open Package 
     File (*.lpk))
  2. Адкрыць zexmlsslib.lpk
  4. Націснуць Усталяваць (Install).
  5. Націснуць "Так"("Yes") калі прапануе перасабраць lazarus.

 2. Delphi 7, BDS 2005, BDS 2006, Delphi 2007, CodeGear RAD Studio 2009/2010, Delphi XE і Delphi XE2.
    Папярэджанне: Delphi 5 будзе працаваць толькі з кадоўкамі 
    windows-1251 і CP866, іншыя кадоўкі прыйдзецца ручкамі (можна 
    паспрабаваць узяць system.pas з Delphi 7)!
  0. Калі не выкарыстоўваеце ZColorStringGrid, дадайце ў src\zexml.inc радок
     {$DEFINE NOZCOLORSTRINGGRID}
  1. Выдаліце папярэднюю версію (адпаведныя bpl, dcp і dcu падчысціце).
  2. Дадайце (калі неабходна) дырэкторыю з зыходным кодам у Tools->
     Environment Options->Library->Library Path (для BDS 2005 ці 
     BDS 2006 у Tools->Options->Environment Options->Delphi Options->
     Library - Win32->Library Path) 
  3. Для выкарыстання функцый ReadXLSX, SaveXmlssToXLSX, SaveXmlssToODFS 
     і ReadODFS: прачытайце /delphizip/readme.txt, заменіце ў дырэкторыі src inc-файлы на /delphizip/патрэбны_пакавальнік
     (пакуль падтрымліваюцца TurboPower Abbrevia (http://sourceforge.net/projects/tpabbrevia/),
     Synzip (http://synopse.info) і JEDI Code library (http://sourceforge.net/projects/jcl/)). Усталюйце патрэбны пакавальнік. 
  4. Адкрыйце патрэбны zexmlsslibe.dpk (ці zexmlsslib.dpk калі не 
     выкарыстоўваецца ZColorStringGrid). Націсніце "Compile" затым "Install".

 3. C++ Builder 6
  0. Калі не выкарыстоўваеце ZColorStringGrid, дадайце ў src\zexml.inc радок
     {$DEFINE NOZCOLORSTRINGGRID}
  1. Выдаліце папярэднюю версію (адпаведныя bpl і bpi падчысціце).
  2. Дадайце (калі неабходна) дырэкторыю з зыходным кодам у 
     Tools->Environment Options->Library->Library Path
  3. Адкрыйце патрэбны zexmlsslibe.bpk (ці zexmlsslib.bpk калі не 
     выкарыстоўваецца ZColorStringGrid). Націсніце "Compile" затым "Install",
     калі праблемы з *.hpp, то вазьміце патрэбныя файлы з \hpp\cbuilderХ\.

Ліцэнзія
--------
Распаўсюджваецца па zlib License.
1. Дадзенае ПА пастаўляецца "як ёсць", аўтар не гарантуе правільную працу 
   софту і не нясе адказнасці за магчымую шкоду ў выніку яго
   выкарыстання.
2. На распаўсюд змененых версій ПА накладваюцца наступныя абмежаванні:
   1. Забараняецца сцвярджаць, што гэта вы напісалі арыгінальны прадукт;
   2. Змененыя версіі не павінны выдавацца за арыгінальны прадукт;
   3. Апавяшчэнне пра ліцэнзію не павінна прыбірацца з пакетаў з зыходнымі тэкстамі.

Папярэджанне!!!
-----------------
У наступных версіях магчыма даданне новых уласцівасцяў/метадаў у класы, 
перанос класаў/працэдур/функцый у іншыя файлы і ўсё такое (бэта ўсёткі).

Кантакты
--------
   Калі ёсць пытанні, каментары ці прапановы то звяртайцеся.
   Аўтар: Небарак Руслан Уладзіміравіч (так у пашпарце напісана)
   url: http:\\avemey.com
   e-mail: support@avemey.com (avemey(мяў).tut.by)