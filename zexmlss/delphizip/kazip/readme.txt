Support KaZip packer by anonym:

after a pair of hours I've added support of "Kazip" opensource library!

A little bit modified version of forked Kazip.pas also included.
Modifications:
 a) replace 'tabulations'->'2spaces';
 b) replace 'string'->'Ansistring', because record TCentralDirectoryFile actually has Ansistring fields.

links to kazip:
1) official:  http://kadao.dir.bg/download/kazip/ - project of September 2005 (seems to be not supported);
2) fork(another author) :  https://github.com/JoseJimeniz/KaZip - "Fixed memory leak " from 21 March 2014, direct link  https://raw.githubusercontent.com/JoseJimeniz/KaZip/master/KaZip.pas

Actual fork of Kazip uses ZlibEx 1.2.8, which may be downloaded from here:   http://www.base2ti.com/

Tested with Delphi XE3 (and quick test for export with Delphi XE).