/'
	Titel: 			csvLibreImport
	File:			csvLibreImport.bas
	Compiler:		FreeBasic
	Editor:         Geany 
	Zeichensatz:	WesteuropÑisch/Westlich (IBM850)
	Zweck:			Importiere csv-Libre-Office Kursdaten
	Wie?:			Du ziehst eine oder mehrere CSV-Dateien auf das
					Icon und das Programm konvertiert und kopiert
					die Daten nach data/csvHist. Die Dateinamen bleiben
					erhalten.	
								
'/

#include "vbcompat.bi"

function GerVal(Zahl as string) as Double	
	dim i as integer
	for i = 1 to len(Zahl)
		if mid(Zahl,i,1) = "," then
			mid(Zahl,i, 1) = "."
		end if
	next i
	return val(Zahl)
end Function

function DatumToDateSerial(Datum as string) as integer
' Achtung bei 2-Stellige Jahreszahlen nur gÅltig fÅr Jahre > 1999
	dim ret as integer
	dim as integer dd, mm, yyyy, p1, i, c
	p1 = 1: c= 1
	for i = p1 to len(Datum)
		' scan for dots and return day, month, year
		if asc(mid(Datum,i,1))=46 then	
			select case c
			case 1	
				dd = val(mid(Datum,p1, i-p1))
				c += 1
				p1 = i+1	
			case 2
				mm = val(mid(Datum,p1, i-p1))				
				p1 = i+1				
			case else
				exit for
			end select
		end if
	next i	
	yyyy = val(mid(Datum,p1))
	if yyyy < 100 then yyyy = yyyy + 2000
	ret = Dateserial(yyyy, mm, dd)
	return ret
end function

function getFileName(FileSpec as string) as string
	dim ret as string
	dim c as integer = len(FileSpec)
	dim i as integer
	for i = c to 1 step -1
		' scan for backslash and return filename
		if asc(mid(FileSpec,i,1))= 92 then
			ret = mid(FileSpec, i+1)
			exit for
		end if
	next i
	return ret
end function

sub csvLibreImport(FileSpec as string)
	dim chIn as integer = freefile
	open FileSpec for input as #chIn
	dim FileOut as string = getFileName(Filespec)
	dim chOut as integer = freefile
	open FileOut for output as #chOut
	dim as string inDate, inOpen, inHigh, inLow, inClose, inTurnover
	dim as integer outDate
	dim as Double outOpen, outHigh, outLow, outClose
	dim as long outTurnover
	input #chIn, inDate, inOpen, inHigh, inLow, inClose, inTurnover
	do while not(eof(chIn))
		input #chIn, inDate, inOpen, inHigh, inLow, inClose, inTurnover
		outDate = DatumToDateSerial(inDate)
		outOpen = GerVal(inOpen)
		outHigh = GerVal(inHigh)
		outLow = GerVal(inLow)
		outClose = GerVal(inClose)
		outTurnover = ValLng(inTurnover)
		write #chOut, outDate, outOpen, outHigh, outLow, outClose, outTurnover
	loop
	
	close chOut, chIn
end sub

' Arbeitsverzeichnis : <drag&drop-Verzeichnis\>csvHist
'        


	dim folder as string = "csvHist"
	dim result as integer = ChDir( folder )

	if result <> 0 then
		' du konntest nicht nach data wechseln erzeuge es
		result = mkDir(folder)
		result = chDir(folder)
		if result <> 0 then
			' Åbel du konntest data nicht erzeugen
			? "Fatal Error : missing folder: data"
			end
		end if
	end if
'---------------------------------------------------------------

	Print Date & " " & Time & " " & Command(0)


	Dim As Integer i = 1
	Do
		Dim As String arg = Command(i)
		If Len(arg) = 0 Then
			Exit Do
		End If

		Print "processing " & i & " = """ & arg & """"
		csvLibreImport(arg)
		i += 1
	Loop

	If i = 1 Then
		Print "(no command line arguments)"
	End If
