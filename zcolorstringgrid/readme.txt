ZColorStringGrid 0.5
Для:
Delphi 6-7, C++Builder 5-6, Borland Developer Studio 2005, BDS 2006, 
CodeGear Delphi 2007, CodeGear RAD Studio 2009 и 2010, XE и XE2.

Наследник TStringgrid. 
Умеет:
 * Выравнивание текста в ячейке по горизонтали и вертикали
 * У каждой ячейки свой "стиль" (цвет фона, шрифт (стиль, цвет), рамка)
 * Возможность объединения ячеек
 * Многострочные ячейки
 * Увеличение размера ячейки, если текст не помещается
 * Поворот многострочного текста

Установка
---------
 1. Delphi 6-7, BDS 2005, BDS 2006, Delphi 2007, CodeGear RAD 
    Studio 2009/2010
    1. Удалите предыдущую версию (соответствующие bpl, dcp и dcu 
       подчистите).
    2. Добавьте (если необходимо) директорию с исходным кодом в Tools->
       Environment Options->Library->Library Path (для BDS 2005 или 
       BDS 2006 в Tools->Options->Environment Options->Delphi Options->
       Library - Win32->Library Path) 
    3. Откройте нужный *.dpk. Нажмите "Compile" затем "Install".

 2. C++ Builder 5-6
    1. Удалите предыдущую версию (соответствующие bpl и bpi подчистите).
    2. Добавьте (если необходимо) директорию с исходным кодом в 
       Tools->Environment Options->Library->Library Path
    3. Откройте нужный zcolor.bpk. Нажмите "Compile" затем "Install",
       если проблемы с *.hpp, то возьмите нужные файлы из \hpp\cbuilderХ\.

 3. Lazarus (только ZCLabel)
    1. Пакет->Открыть файл	пакета (*.lpk) (Components->Open Package 
     File (*.lpk))
    2. Открыть zclabel.lpk
    3. Нажать Установить (Install).
    4. Нажать "Да"("Yes") когда предложит пересобрать lazarus.

Контакты
--------
	Если есть вопросы, комментарии или предложения то обращайтесь.
	Неборак Руслан
	url: http:\\avemey.com
	e-mail: support@avemey.com