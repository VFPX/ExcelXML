*/ Converts a cursor in a Excel XML file without a Grid control

create cursor students ( name c(30), birthday d null, age n(3) )
insert into students values ( "Brian",     date(1500,01,01), 7 )
insert into students values ( "Megan",     date(), 0 )
insert into students values ( "Melanie",   date(1945,03,25), 59 )
insert into students values ( "Stephanie", date(1978,05,24), 35 )
insert into students values ( "Angelina",  date(2011,06,19), 2 )
insert into students values ( "Richard",   date(1995,01,13), 13 )
insert into students values ( "Michael",   date(1982,03,24), 31 )
insert into students values ( "Ingrid",    date(2005,11,18), 7 )
insert into students values ( "Michelle",  date(1978,12,15), 34 )
insert into students values ( "Ryan",  	   date(1999,09,05), 14 )
insert into students values ( "Brian",     date(2005,11,27), 7 )
insert into students values ( "Megan",     date(2001,03,30), 12 )
insert into students values ( "Melanie",   date(1954,02,28), 59 )
insert into students values ( "Stephanie", date(1978,05,24), 35 )
insert into students values ( "Angelina",  date(2011,06,19), 2 )
insert into students values ( "Richard",   date(1995,01,13), 13 )
insert into students values ( "Michael",   date(1982,03,24), 31 )
insert into students values ( "Ingrid",    date(2005,11,18), 7 )
insert into students values ( "Michelle",  date(1978,12,15), 34 )
insert into students values ( "Ryan",  	   date(1999,09,05), 14 )

local loExcelXML, llOk
loExcelXML = NewObject("ExcelXML","..\ExcelXML.prg")
loExcelXML.SheetName = "Students Grade 5A"
loExcelXML.OpenAfterSaving = .t.
llOk = loExcelXML.Save("Sample05.XML")




if llOk
	messagebox("File saved", 64)
else
	messagebox("File not saved", 16)
endif

use in students
