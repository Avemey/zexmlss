ZEXMLSSLIB 0.0.14 (alpha)
��� Lazarus, Delphi 7, C++Builder 6.
Borland Developer Studio 2005, BDS 2006, CodeGear Delphi 2007, 
CodeGear RAD Studio 2009, 2010, Delphi XE � Delphi XE2.
(Delphi 5 � 6 �� ��������. Delphi 5 ������ � ����������� windows-1251 � CP866, �� ����� ��������
���� �������)

������ � ������� Excel XML Spreadsheet / OpenDocument Format (ODS) / OOXML (xlsx) ��� �������������� Excel-�/OO Calc-�.

���������
---------
 1. Lazarus
  0. ���� �� ����������� ������ ������������� LConvEncoding.pas, ������� 
     � src\zexml.inc ������ {$DEFINE USELCONVENCODING} (� UTF-8 ����� 
     �������� ���������, � ��������� ��������� ������� �������)
  1. �����->������� ����	������ (*.lpk) (Components->Open Package 
     File (*.lpk))
  2. ������� zexmlsslib.lpk
  4. ������ ���������� (Install).
  5. ������ "��"("Yes") ����� ��������� ����������� lazarus.

 2. Delphi 5-7, BDS 2005, BDS 2006, Delphi 2007, CodeGear RAD Studio 2009/2010, Delphi XE � Delphi XE2.
    ��������������: Delphi 5 ����� �������� ������ � ����������� 
    windows-1251 � CP866, ������ ��������� ������� ������� (����� 
    ����������� ����� system.pas �� Delphi 7)!
  0. ���� �� ����������� ZColorStringGrid, �������� � src\zexml.inc ������
     {$DEFINE NOZCOLORSTRINGGRID}
  1. ������� ���������� ������ (��������������� bpl, dcp � dcu ����������).
  2. �������� (���� ����������) ���������� � �������� ����� � Tools->
     Environment Options->Library->Library Path (��� BDS 2005 ��� 
     BDS 2006 � Tools->Options->Environment Options->Delphi Options->
     Library - Win32->Library Path) 
  3. ��� ������������� ������� ReadXLSX, SaveXmlssToXLSX, SaveXmlssToODFS 
     � ReadODFS: ���������� /delphizip/readme.txt, �������� � ����� src inc-����� �� /delphizip/������_���������
     (���� �������������� TurboPower Abbrevia (http://sourceforge.net/projects/tpabbrevia/),
     Synzip (http://synopse.info) � JEDI Code library (http://sourceforge.net/projects/jcl/)). 
     ���������� ������ ���������.
  4. �������� ������ zexmlsslibe.dpk (��� zexmlsslib.dpk ���� �� 
     ������������ ZColorStringGrid). ������� "Compile" ����� "Install".

 3. C++ Builder 6
  0. ���� �� ����������� ZColorStringGrid, �������� � src\zexml.inc ������
     {$DEFINE NOZCOLORSTRINGGRID}
  1. ������� ���������� ������ (��������������� bpl � bpi ����������).
  2. �������� (���� ����������) ���������� � �������� ����� � 
     Tools->Environment Options->Library->Library Path
  3. �������� ������ zexmlsslibe.bpk (��� zexmlsslib.bpk ���� �� 
     ������������ ZColorStringGrid). ������� "Compile" ����� "Install",
     ���� �������� � *.hpp, �� �������� ������ ����� �� \hpp\cbuilder�\.

��������
--------
���������������� �� zlib License.
1. ������ �� ������������ "��� ����", ����� �� ����������� ���������� ������ 
   ����� � �� ����� ��������������� �� ��������� ����� � ���������� ���
   �������������.
2. �� ��������������� ��������� ������ �� ������������� ��������� �����������:
   1. ����������� ����������, ��� ��� �� �������� ������������ �������;
   2. ��������� ������ �� ������ ���������� �� ������������ �������;
   3. ����������� � �������� �� ������ ��������� �� ������� � ��������� ��������.

��������������!!!
-----------------
� ��������� ������� �������� ���������� ����� �������/������� � ������, 
������� �������/��������/������� � ������ ����� � �� ����� (���� ��-����).

��������
--------
   ���� ���� �������, ����������� ��� ����������� �� �����������.
   �����: ������� ������ ������������
   url: http:\\avemey.com
   e-mail: support@avemey.com (avemey(���).tut.by)