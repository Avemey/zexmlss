ZEXMLSSLIB 0.0.15 (beta)
Для Lazarus, Delphi 7, C++Builder 6.
Borland Developer Studio 2005, BDS 2006, CodeGear Delphi 2007, 
CodeGear RAD Studio 2009, 2010, Delphi XE и Delphi XE2.
(Delphi 5 И 6 не проверял. Delphi 5 только с кодировками windows-1251 и CP866, но можно добавить
свои функции)

Работа с файлами Excel XML Spreadsheet / OpenDocument Format (ODS) / OOXML (xlsx) без установленного Excel-а/OO Calc-а.

Установка
---------
 1. Lazarus
  0. Если не используете модуль перекодировки LConvEncoding.pas, уберите 
     в src\zexml.inc строку {$DEFINE USELCONVENCODING} (с UTF-8 будет 
     работать нормально, а остальные кодировки придётся ручками)
  1. Пакет->Открыть файл	пакета (*.lpk) (Components->Open Package 
     File (*.lpk))
  2. Открыть zexmlsslib.lpk
  4. Нажать Установить (Install).
  5. Нажать "Да"("Yes") когда предложит пересобрать lazarus.

 2. Delphi 5-7, BDS 2005, BDS 2006, Delphi 2007, CodeGear RAD Studio 2009/2010, Delphi XE и Delphi XE2.
    Предупреждение: Delphi 5 будет работать только с кодировками 
    windows-1251 и CP866, другие кодировки придётся ручками (можно 
    попробовать взять system.pas из Delphi 7)!
  0. Если не используете ZColorStringGrid, добавьте в src\zexml.inc строку
     {$DEFINE NOZCOLORSTRINGGRID}
  1. Удалите предыдущую версию (соответствующие bpl, dcp и dcu подчистите).
  2. Добавьте (если необходимо) директорию с исходным кодом в Tools->
     Environment Options->Library->Library Path (для BDS 2005 или 
     BDS 2006 в Tools->Options->Environment Options->Delphi Options->
     Library - Win32->Library Path) 
  3. Для использования функций ReadXLSX, SaveXmlssToXLSX, SaveXmlssToODFS 
     и ReadODFS: прочитайте /delphizip/readme.txt, замените в папке src inc-файлы на /delphizip/нужный_упаковщик
     (пока поддерживаются TurboPower Abbrevia (http://sourceforge.net/projects/tpabbrevia/),
     Synzip (http://synopse.info) и JEDI Code library (http://sourceforge.net/projects/jcl/)). 
     Установите нужный упаковщик.
  4. Откройте нужный zexmlsslibe.dpk (или zexmlsslib.dpk если не 
     используется ZColorStringGrid). Нажмите "Compile" затем "Install".

 3. C++ Builder 6
  0. Если не используете ZColorStringGrid, добавьте в src\zexml.inc строку
     {$DEFINE NOZCOLORSTRINGGRID}
  1. Удалите предыдущую версию (соответствующие bpl и bpi подчистите).
  2. Добавьте (если необходимо) директорию с исходным кодом в 
     Tools->Environment Options->Library->Library Path
  3. Откройте нужный zexmlsslibe.bpk (или zexmlsslib.bpk если не 
     используется ZColorStringGrid). Нажмите "Compile" затем "Install",
     если проблемы с *.hpp, то возьмите нужные файлы из \hpp\cbuilderХ\.

Лицензия
--------
Распространяется по zlib License.
1. Данное ПО поставляется "как есть", автор не гарантирует правильную работу 
   софта и не несет ответственности за возможный ущерб в результате его
   использования.
2. На распространение изменённых версий ПО накладываются следующие ограничения:
   1. Запрещается утверждать, что это вы написали оригинальный продукт;
   2. Изменённые версии не должны выдаваться за оригинальный продукт;
   3. Уведомление о лицензии не должно убираться из пакетов с исходными текстами.

Предупреждение!!!
-----------------
В следующих версиях возможно добавление новых свойств/методов в классы, 
перенос классов/процедур/функций в другие файлы и всё такое (бета всё-таки).

Контакты
--------
   Если есть вопросы, комментарии или предложения то обращайтесь.
   Автор: Неборак Руслан Владимирович
   url: http:\\avemey.com
   e-mail: support@avemey.com (avemey(мяу).tut.by)