* ToDo:
* The new SET TEXTMERGE TO MEMVAR option is not detected



Lparameters tlRun
Local lcProgram, lcSource, loParser

_vfp.AutoYield = .t.

If tlRun
	* Create the parser object
	loParser = CreateObject('ParseLocals')
	* Try to load the source from the current editing window
	lcSource = loParser.LoadSource()
	
	If !Empty(lcSource)
		* Set the initial function name to the window title
		loParser.cFunction = Proper(Wontop())
		* Remove any read only markers
		loParser.cFunction = Trim(Strtran(loParser.cFunction, '[read Only]'))
		* Remove any old error files
		loParser.RemoveErrorFile()
		* Analyze the source
		loParser.Execute(lcSource)
		* Display the errors of any where present
		loParser.ShowErrorFile()
	EndIf
Else
	* Register the option in the tools menu
	Define Bar 991 of _mtools Prompt '\-'
	Define Bar 992 of _mtools Prompt 'Check \<local declarations' Key Alt+L, 'Alt+L'
	lcProgram = '"'+ Sys(16) + '"'
	On Selection Bar 992 of _mtools do &lcProgram With .t.
EndIf

Return



#Define c_System        1
#Define c_Parameter     2
#Define c_Variable      3
#Define c_Array         4

#Define c_crlf          Chr(13)+Chr(10)
#Define c_tab           Chr(9)


********************************
Define Class ParseLocals as Line
********************************

Dimension aDeclaredVar[1, 2]
Dimension aAssignedVar[1, 2]

cClass    = ''
cFunction = ''

cErrorFile = ''

oRegExp    = Null
oTherm      = Null


**********************
Procedure LoadSource()
**********************
* Load all the source form the current editing window
Local array laEdEnv[25]
Local lcSource, lhWonTop, lnEdPos, lcClipText
lcSource = ''
laEdEnv = ''
* We need to load the FoxTools library for the _ed... editor functions
Set Library To Home()+'foxtools' Additive
* Get a handle on the topmost window
lhWonTop = _WOnTop()

* Retreive the current editor information
_EdGetEnv(lhWonTop, @laEdEnv)
If laEdEnv[25] <> 0
	* We aren't in the command window

	* Save the current clipboard contents
	lcClipText = _ClipText
	_ClipText  = ''
	* Copy the selected code onto the clipboard
	_EdCopy(lhWonTop)
	If Empty(_ClipText)
		* Save the current editing posistion
		lnEdPos = _EdGetPos(lhWonTop)
		* Select all the code in the editor
		_EdSelect(lhWonTop, 0, 999999)
		* Copy the code onto the clipboard
		_EdCopy(lhWonTop)
		* Reset the previous editing position
		_EdSetPos(lhWonTop, lnEdPos)
		* Make sure the cursor is visible
		_EdSToSel(lhWonTop, .t.)
	EndIf
	* Retreive the code from the clipboard
	lcSource = _Cliptext
	* Restore the clipboard
	_Cliptext = lcClipText
EndIf

Return lcSource


***************************
Procedure Execute(tcSource)
***************************
* Split into classes and methods and parse them
Local loRegExp as VbScript.RegExp
Local lcSource, lnCount, loMatches, lnStart, lnEnd, lcTemp, lnI

* Display the progress in a thermometer
This.oTherm = CreateObject('Thermometer', 'Parsing source code')
This.oTherm.Show()
This.oTherm.Update(0, 'Cleaning source')

* Do a cleanup to remove comments and tab charaters
lcSource = This.CleanupSource(tcSource)
* Create a regular expressions object
loRegExp = This.GetRegExp()
* Set the patern to split the code into separate classes and functions
loRegExp.Pattern = '(^Define +Class)|(^Proc(?:edure)?)|(^Func(?:tion)?)|(^EndDefine)'
loMatches = loRegExp.Execute(lcSource)
lnCount = loMatches.Count

* Loop through all the separate class and function code snippets
lnStart = 1
For lnI = 0 to lnCount - 1
	* Determine the end of the code snippet
	lnEnd = loMatches.Item(lnI).FirstIndex
	* Extract the code snippet
	lcTemp = Substr(lcSource, lnStart, lnEnd - lnStart)
	* And parse it
	This.CheckSection(lcTemp)
	lnStart = lnEnd
	* Display the progress
	This.oTherm.Update(100 * (lnI + 1) / (lnCount - 1), This.cClass - '.' - This.cFunction)
EndFor

* Extract the remainder of the code or all the code if there is no function/define class statement
lnEnd = Len(lcSource)
lcTemp = Substr(lcSource, lnStart, lnEnd - lnStart)
* And parse it
This.CheckSection(lcTemp)


* Display the fact that we are finished
This.oTherm.Update(100, 'Done')

Return


*********************************
Procedure CheckSection(tcSection)
*********************************
* Analyze an indiviual code snippet. This assumes the code has been split into separate functions

Local array laLines[1], laTemp[1]
Local lcSection, lnCount, lnI, lcLine, llParse, lnAtPos, lcVar, lcExact
lcSection = tcSection
lcExact   = Set("Exact")
Set Exact Off

DO While Left(lcSection, 1) = Chr(13) or Left(lcSection, 1) = Chr(10)
	lcSection = Substr(lcSection, 2)
EndDo

This.Reset()

lnCount = ALines(laLines, lcSection)
lcLine = laLines[1]

llParse = .f.

DO case
Case lnCount = 0
	* No data to check
Case InList(Proper(lcLine), 'Function ', 'Func ')
	lnAtPos = At('(', lcLine)
	If lnAtPos > 0
		lcLine = Left(lcLine, lnAtPos - 1)
	EndIf
	lcLine = Substr(lcLine, At(' ', lcLine))
	This.cFunction = Alltrim(lcLine)
	llParse = .t.
Case InList(Proper(lcLine), 'Procedure ', 'Proc ')
	lnAtPos = At('(', lcLine)
	If lnAtPos > 0
		lcLine = Left(lcLine, lnAtPos - 1)
	EndIf
	lcLine = Substr(lcLine, At(' ', lcLine))
	This.cFunction = Alltrim(lcLine)
	llParse = .t.
Case Proper(lcLine) = 'Define ' and ' Class ' $ Proper(lcLine)
	lnAtPos = Atc('Class ', lcLine)
	lcLine = Alltrim(Substr(lcLine, lnAtPos + 6))
	lnAtPos = At(' ', lcLine)
	If lnAtPos > 0
		lcLine = Left(lcLine, lnAtPos - 1)
	EndIf
	This.cClass = lcLine
	This.cFunction = ''
Case Proper(lcLine) + ' ' = 'EndDefine '
	This.cClass = ''
	This.cFunction = ''
Otherwise
	llParse = .t.
	lnAtPos = 0
EndCase

Set Exact &lcExact

If llParse
	If lnAtPos > 0
		lcLine = laLines[1]
		lcLine = Alltrim(Substr(lcLine, lnAtPos + 1))
		lnAtPos = At(')', lcLine)
		If lnAtPos > 0
			lcLine = Alltrim(Left(lcLine, lnAtPos - 1))
		EndIf
		
		* Parameters found
		lnCount = ALines(laTemp, lcLine, .t., ',')
		For lnI = 1 to lnCount
			lcVar = laTemp[lnI]
			This.AddDeclaredVar(lcVar, c_Parameter)
		EndFor
	EndIf

	This.ParseSection(lcSection)
	This.ReportSection()
EndIf


Return


*************************
Procedure ReportSection()
*************************
Local lnCount, lnI, lnFound, lcVar, lcMessage, lcTemp, lnStart, lnAtPos
lcMessage = ''

lnCount = Alen(This.aAssignedVar, 1)
If Empty(This.aAssignedVar[1])
	lnCount = 0
endif

For lnI = lnCount To 1 Step -1
	lcVar = This.aAssignedVar[lnI, 1]
	lnFound = Ascan(This.aDeclaredVar, lcVar, -1, -1, -1, 1+2+4+8)
	If lnFound > 0
		Adel(This.aAssignedVar, lnI)
		Adel(This.aDeclaredVar, lnFound)
	EndIf
EndFor


lcTemp = ''
For lnI = 1 To Alen(This.aAssignedVar, 1)
	DO Case
	Case Empty(This.aAssignedVar[lnI, 1]) 
		* No memory variable found
	Case This.aAssignedVar[lnI, 2] = c_System
		* This is a VFP internal, ignore it
	Case This.aAssignedVar[lnI, 2] = c_Array
		lcTemp = lcTemp + c_tab + 'Local Array ' + This.aAssignedVar[lnI, 1] + '[1]' + c_crlf
	Otherwise
	EndCase
EndFor

If !Empty(lcTemp)
	lcMessage = lcMessage +  'Add ;' + c_crlf + lcTemp + c_crlf
endif

lcTemp = ''
For lnI = 1 To Alen(This.aAssignedVar, 1)
	DO Case
	Case Empty(This.aAssignedVar[lnI, 1]) 
		* No memory variable found
	Case This.aAssignedVar[lnI, 2] = c_System
		* This is a VFP internal, ignore it
	Case This.aAssignedVar[lnI, 2] = c_Array
	Otherwise
		lcTemp = lcTemp + This.aAssignedVar[lnI, 1] + ', '
	EndCase
EndFor


If !Empty(lcTemp)
	lcTemp = Left(lctemp, Len(lcTemp) - 2)

	lnStart = 0
	For lnI = 1 To Occurs(',', lcTemp)
		lnAtPos = At(',', lcTemp, lnI) - lnStart
		If lnAtPos - lnStart > 80
			lnStart = lnAtPos 
			lcTemp = Stuff(lcTemp, lnAtPos, 2, c_crlf)
		EndIf
	EndFor

	lcTemp = Strtran(lcTemp, c_crlf , c_crlf + c_tab + 'Local ')
	lcMessage = lcMessage +  'Add ;' + c_crlf + c_tab + 'Local ' + lcTemp + c_crlf
endif

lcTemp = ''
For lnI = 1 To Alen(This.aDeclaredVar, 1)
	DO Case
	Case Empty(This.aDeclaredVar[lnI, 1])
	Case This.aDeclaredVar[lnI, 2] = c_Parameter
	Case This.aDeclaredVar[lnI, 2] = c_Array
		lcTemp = lcTemp + 'Array ' + This.aDeclaredVar[lnI, 1] + '[1], '
	Otherwise
		lcTemp = lcTemp + This.aDeclaredVar[lnI, 1] + ', '
	EndCase
EndFor

If !Empty(lcTemp)
	lcTemp = Left(lctemp, Len(lcTemp) - 1)
	lcTemp = Strtran(lcTemp, ',', ', ;' + c_crlf + c_tab)
	lcMessage = lcMessage +  c_crlf + 'Remove Local ;' + c_crlf + c_tab + lcTemp + c_crlf
EndIf

If !Empty(lcMessage)
	If Empty(This.cClass)
		lcTemp = 'Function ' + This.cFunction
	Else
		lcTemp = 'Class ' + This.cClass + ' - Method ' + This.cFunction
	EndIf
	lcTemp = Replicate('*', Len(lcTemp)) + c_crlf + lcTemp + c_crlf + Replicate('*', Len(lcTemp))
	
	lcMessage = lcTemp + c_crlf + lcMessage + c_crlf
	StrToFile(lcMessage, This.cErrorFile, 1)
EndIf

Return




*********************************
Procedure ParseSection(lcSection)
*********************************
Local Array laTemp1[1]
Local Array laTemp2[1]
Local loRegExp as VbScript.RegExp
Local loMatches, lnCount, lnI, lcVar, lnVar, lcFunction, lcTemp, lcPrefix
loRegExp = This.GetRegExp()

* Find all declared variables
loRegExp.Pattern = '^(Local|Private|Public|L?Parameters?) +(.+?)$'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcTemp = Lower(loMatches.Item(lnI).SubMatches(0))
	lcVar = loMatches.Item(lnI).SubMatches(1)
	lnCount = ALines(laTemp1, lcVar, .t., ',')
	For lnVar = 1 to lnCount
		lcVar = laTemp1[lnVar]

		DO Case
		Case InList(Atc('param', loMatches.Item(lnI).SubMatches(0)), 1, 2)
			This.AddDeclaredVar(lcVar, c_Parameter)
		Case lcTemp == 'private' and Lower(lcVar) == 'all'
			* A PRIVATE ALL statement, this is bad practice, ignore it
		Otherwise
			This.AddDeclaredVar(lcVar, c_Variable)
		EndCase
	EndFor
EndFor

* Find assignments using the =
loRegExp.Pattern = '^(?:m\.)?(\w+) *(\[.*?)?='
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcVar = loMatches.Item(lnI).SubMatches(0)
	DO case
	Case !IsNull(loMatches.Item(lnI).SubMatches(1))
		* Found part of an array reference
		This.AddAssignedVar(lcVar, c_Array)
	Otherwise
		* Just a regular assignment
		This.AddAssignedVar(lcVar, c_Variable)
	EndCase
EndFor

* Add the FOR and FOR EACH assignments
loRegExp.Pattern = '^For +(?:Each +)?(?:m\.)?(\w*)'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcVar = loMatches.Item(lnI).SubMatches(0)
	* Just a regular assignment
	This.AddAssignedVar(lcVar, c_Variable)
EndFor

* Find assignments using STORE
loRegExp.Pattern = '^store\s+.*?\s+to\s+(.+?)$'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcVar = loMatches.Item(lnI).SubMatches(0)
	lnCount = ALines(laTemp1, lcVar, .t., ',')
	For lnVar = 1 to lnCount
		lcVar = laTemp1[lnVar]
		DO Case
		Case At('.', lcVar) > 0 and Lower(Left(lcVar, 2)) <> 'm.'
			* This is an asignment of an object property
		Case Left(lcVar, 1) $ '(&'
			* Macro or named reference, ignore
		Otherwise
			This.AddAssignedVar(lcVar, c_Variable)
		EndCase
	EndFor
EndFor

* Find assignments using SCATTER
loRegExp.Pattern = '^scatter +(?:memo  +)?name +(\w+)'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcVar = loMatches.Item(lnI).SubMatches(0)
	This.AddAssignedVar(lcVar, c_Variable)
EndFor

* Find all (IN)TO ARRAY creations
loRegExp.Pattern = 'to array (\w+)$'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcVar = loMatches.Item(lnI).SubMatches(0)
	This.AddAssignedVar(lcVar, c_Array)
EndFor

* Find all TO creations
loRegExp.Pattern = '(?:Calcutate|Sum|Text|Catch) to (?:m\.)?(\w+)'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcVar = loMatches.Item(lnI).SubMatches(0)
	This.AddAssignedVar(lcVar, c_Variable)
EndFor

* 2007-02-07 Adde on request of Peter Crabtree
* Find all DO FORM NAME creations
loRegExp.Pattern = '^do form.* NAME (?:m\.)?(\w+).*$'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
	lcVar = loMatches.Item(lnI).SubMatches(0)
	This.AddAssignedVar(lcVar, c_Variable)
ENDFOR

* 2007-02-07 Added on request of Kurt Gra�l
* Find vars using m.-notation
loRegExp.Pattern = '([a-zA-Z0-9_]*m\.)([a-zA-Z0-9_]*)'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
For lnI = 0 to lnCount - 1
     lcVar = loMatches.Item(lnI).SubMatches(1)
     lcPrefix = Upper(loMatches.Item(lnI).SubMatches(0))
     If Upper(loMatches.Item(lnI).SubMatches(0)) == "M."
         This.AddAssignedVar(lcVar, c_Variable)
     EndIf
EndFor


* Find all AFunction array assignments
loRegExp.Pattern = '\b(a\w*)\((\w*)\b[ ,)].*?$'
loMatches = loRegExp.Execute(lcSection)
lnCount = loMatches.Count
* Fill an array with all native functions
ALanguage(laTemp1, 2)

For lnI = 0 to lnCount - 1
	lcFunction = Lower(loMatches.Item(lnI).SubMatches(0))
	lcVar = loMatches.Item(lnI).SubMatches(1)
	DO Case 
	Case Ascan(laTemp1, lcFunction, -1, -1, -1,1+2+4+8) = 0
		* The function isn't a native VFP function
	Case InList(lcFunction, 'alen', 'adel', 'ascan', 'asort', 'asubscript')
		* These array functions don't create an array
	Case InList(lcFunction, 'addbs', 'alltrim', 'asc', 'at', 'at_c', 'atc', 'atcc', 'atcline', 'atline')
		* These string functions don't create an array
	Case InList(lcFunction, 'abs', 'acos', 'asin', 'atan', 'atn2')
		* These math functions don't create an array
	Case InList(lcFunction, 'alias')
		* These xbase functions don't create an array
	Case InList(lcFunction, 'acopy')
		* These ACopy create the second parameter
		Activate Screen
		lcTemp = loMatches.Item(lnI).Value
		ALines(laTemp2, lcTemp, .t., ',', ')')
		lcVar = laTemp2[2]
		If At('.', lcVar) = 0
			* Add it the the assignments
			This.AddAssignedVar(lcVar, c_Array)
		EndIf
	Otherwise
		* Add it the the assignments
		This.AddAssignedVar(lcVar, c_Array)
	EndCase
EndFor

Return


***************************************
Procedure AddDeclaredVar(tcVar, tnType)
***************************************
* Add a variable declaration
Local lcVar, lnAtPos, lnSize
lcVar = Alltrim(tcVar)

If Left(Upper(lcVar), 5) = 'ARRAY'
	* Its an array, strip the array from the name
	lcVar = Alltrim(Substr(lcVar, 7))
	tnType = Max(tnType, c_Array)
EndIf

lnAtPos = At('[', lcVar)
If lnAtPos > 0
	* Square brackets indicate an array
	lcVar = Left(lcVar, lnAtPos - 1)	
	tnType = Max(tnType, c_Array)
EndIf

lnAtPos = At('(', lcVar)
If lnAtPos > 0
	* Square brackets indicate an array
	lcVar = Left(lcVar, lnAtPos - 1)	
	tnType = Max(tnType, c_Array)
EndIf

If Upper(Left(lcVar, 2)) = 'M.'
	* Strip any addional m. prefix
	lcVar = Substr(lcVar, 3)
EndIf

lnAtPos = Atc(' as ', lcVar)
If lnAtPos > 0
	* We don't need the variable type either
	lcVar = Trim(Left(lcVar, lnAtPos))
EndIf

If !IsAlpha(lcVar)
	* Only save variables starting with a letter
	lcVar = ''
EndIf

If !Empty(lcVar) And Ascan(this.aDeclaredVar, lcVar,-1, -1, -1, 1+2+4+8) = 0
	* A new varaible, save it
	lnSize = Alen(This.aDeclaredVar, 1)
	If lnSize > 1 or !Empty(This.aDeclaredVar)
		lnSize = lnSize + 1 
		Dimension This.aDeclaredVar[lnSize, 2]
	EndIf

	This.aDeclaredVar[lnSize, 1] = lcVar
	This.aDeclaredVar[lnSize, 2] = tnType
EndIf

Return


***************************************
Procedure AddAssignedVar(tcVar, tnType)
***************************************
* Add a variable assignment
Local lcVar, lnRow

lcVar = Alltrim(tcVar)
If Left(lcVar, 1) = '_'
	* Variables starting with an _ are supposed to be VFP internal
	tnType = c_System
EndIf

* Locate the variable
lnRow = Ascan(this.aAssignedVar, lcVar,-1, -1, -1,1+2+4+8)
If lnRow = 0
	* Not found, add it
	lnRow = Alen(This.aAssignedVar, 1)
	If lnRow > 1 or !Empty(This.aAssignedVar)
		lnRow = lnRow + 1 
		Dimension This.aAssignedVar[lnRow, 2]
	EndIf
	* Add the variable name and its type
	This.aAssignedVar[lnRow, 1] = lcVar
	This.aAssignedVar[lnRow, 2] = tnType
Else
	* Update the type if needed
	This.aAssignedVar[lnRow, 2] = Max(This.aAssignedVar[lnRow, 2], tnType)
EndIf

Return


*********************
Procedure GetRegExp()
*********************
* Return the Regular Expression obj

If IsNull(This.oRegExp)
	* First time
	* Create a Regular Expression object
	This.oRegExp = CreateObject('VbScript.RegExp')
EndIf

* Set the default properties
This.oRegExp.Multiline=.t.
This.oRegExp.Global=.t.
This.oRegExp.IgnoreCase = .t.

Return This.oRegExp


*****************
Procedure Reset()
*****************
* Reset the internal memory so we can start a new procedure
Dimension This.aDeclaredVar[1, 2]
This.aDeclaredVar[1, 1] = ''
This.aDeclaredVar[1, 2] = 0

Dimension This.aAssignedVar[1, 2]
This.aAssignedVar[1, 1] = ''
This.aAssignedVar[1, 2] = 0

Return


*********************************
Procedure CleanupSource(tcSource)
*********************************
* Remove comments and special characters
Local array laLines[1]
Local lcSource, lnCount, lnI, lcLine, lnAtPos
lcSource = tcSource

* Remove all linefeeds
* Change tabs into spaces
lcSource = Chrtran(lcSource, Chr(9)+Chr(10), ' ')

lnCount = ALines(laLines, lcSource, .t.)
* Remove all comments
For lnI = 1 to lnCount
	lcLine = Alltrim(laLines[lnI])
	DO case
	Case Lower(Left(lcLine, 7)) = '*:local'
		* Treat as a normal local declaration
		lcLine = Substr(lcLine, 3)
	Case Lower(Left(lcLine, 10)) = 'protected '
		* Remove the Protected attribute
		lcLine = Alltrim(Substr(lcLine, 10))
	Case Lower(Left(lcLine, 7)) = 'hidden '
		* Remove the Hidden attribute
		lcLine = Alltrim(Substr(lcLine, 7))
	Case Left(lcLine, 1) = '*'
		* Just a comment, ignore
		lcLine = ''
	EndCase
	* Remove all inline comments
	lnAtPos = At('&'+ '&', lcLine)
	If lnAtPos > 0
		lcLine = Trim(Left(lcLine, lnAtPos - 1))
	EndIf
	laLines[lnI] = lcLine
EndFor

* Turn all multiline statements into a single line
For lnI = lnCount - 1 to 1 Step -1
	lcLine = laLines[lnI]
	If Right(lcLine, 1) = ';'
		* This is a multi line statement, move to a single line
		lcLine = Left(lcLine, Len(lcLine) - 1)
		lcLine = lcLine + laLines[lnI + 1]
		laLines[lnI + 1] = ''
		laLines[lnI] = lcLine
	EndIf
EndFor
	
* Remove the contents of TEXT TO Blocks as they can confuse the parser
For lnI = 1 to lnCount
	lcLine = Alltrim(laLines[lnI])
	If Lower(Left(lcLine, 5)) = 'text '
		* Found the start of a TEXT block
		DO while lnI < lnCount
			* Remove the contents untill the EndText
			lnI = lnI + 1
			lcLine = Alltrim(laLines[lnI])
			If Lower(Left(lcLine, 7)) = 'endtext'
				* Found the EndText, continue and search for another block
				Exit
			Else
				* Part of the Text/EndText block, strip it
				laLines[lnI] = ''
			Endif
		EndDo 
	EndIf
EndFor

lcSource = Chr(13)
For each lcLine in laLines 
	lcSource = lcSource + lcLine + Chr(13)
EndFor

Return lcSource


***************************
Procedure RemoveErrorFile()
***************************

This.cErrorFile = FullPath(ForceExt(Program(), 'err'))
	
If Wexist(JustFname(This.cErrorFile))
	Release Windows (JustFname(This.cErrorFile))
endif
Erase (This.cErrorFile)
Return


*************************
Procedure ShowErrorFile()
*************************

If File(This.cErrorFile)
	Modify Comm (This.cErrorFile) NoWait 
EndIf

Return

******************************************
Procedure Error(tnError, tcMethod, tnLine)
******************************************
* Ignore all runtime errors for developer convenience

Return


EndDefine






* The remainder of the code is the thermometer from the standard vfpxtab source code
#DEFINE WIN32FONT			'MS Sans Serif'

***********************************************************************
***********************************************************************
DEFINE CLASS thermometer AS form

	Top = 196
	Left = 142
	Height = 88
	Width = 356
	AutoCenter = .T.
	BackColor = RGB(192,192,192)
	BorderStyle = 0
	Caption = ""
	Closable = .F.
	ControlBox = .F.
	MaxButton = .F.
	MinButton = .F.
	Movable = .F.
	AlwaysOnTop = .F.
	ipercentage = 0
	ccurrenttask = ''
	shpthermbarmaxwidth = 322
	cthermref = ""
	Name = "thermometer"

	ADD OBJECT shape10 AS shape WITH ;
		BorderColor = RGB(128,128,128), ;
		Height = 81, ;
		Left = 3, ;
		Top = 3, ;
		Width = 1, ;
		Name = "Shape10"


	ADD OBJECT shape9 AS shape WITH ;
		BorderColor = RGB(128,128,128), ;
		Height = 1, ;
		Left = 3, ;
		Top = 3, ;
		Width = 349, ;
		Name = "Shape9"


	ADD OBJECT shape8 AS shape WITH ;
		BorderColor = RGB(255,255,255), ;
		Height = 82, ;
		Left = 352, ;
		Top = 3, ;
		Width = 1, ;
		Name = "Shape8"


	ADD OBJECT shape7 AS shape WITH ;
		BorderColor = RGB(255,255,255), ;
		Height = 1, ;
		Left = 3, ;
		Top = 84, ;
		Width = 350, ;
		Name = "Shape7"


	ADD OBJECT shape6 AS shape WITH ;
		BorderColor = RGB(128,128,128), ;
		Height = 86, ;
		Left = 354, ;
		Top = 1, ;
		Width = 1, ;
		Name = "Shape6"


	ADD OBJECT shape4 AS shape WITH ;
		BorderColor = RGB(128,128,128), ;
		Height = 1, ;
		Left = 1, ;
		Top = 86, ;
		Width = 354, ;
		Name = "Shape4"


	ADD OBJECT shape3 AS shape WITH ;
		BorderColor = RGB(255,255,255), ;
		Height = 85, ;
		Left = 1, ;
		Top = 1, ;
		Width = 1, ;
		Name = "Shape3"


	ADD OBJECT shape2 AS shape WITH ;
		BorderColor = RGB(255,255,255), ;
		Height = 1, ;
		Left = 1, ;
		Top = 1, ;
		Width = 353, ;
		Name = "Shape2"


	ADD OBJECT shape1 AS shape WITH ;
		BackStyle = 0, ;
		Height = 88, ;
		Left = 0, ;
		Top = 0, ;
		Width = 356, ;
		Name = "Shape1"


	ADD OBJECT shape5 AS shape WITH ;
		BorderStyle = 0, ;
		FillColor = RGB(192,192,192), ;
		FillStyle = 0, ;
		Height = 15, ;
		Left = 17, ;
		Top = 47, ;
		Width = 322, ;
		Name = "Shape5"


	ADD OBJECT lbltitle AS label WITH ;
		FontName = WIN32FONT, ;
		FontSize = 8, ;
		BackStyle = 0, ;
		BackColor = RGB(192,192,192), ;
		Caption = "", ;
		Height = 16, ;
		Left = 18, ;
		Top = 14, ;
		Width = 319, ;
		WordWrap = .F., ;
		Name = "lblTitle"


	ADD OBJECT lbltask AS label WITH ;
		FontName = WIN32FONT, ;
		FontSize = 8, ;
		BackStyle = 0, ;
		BackColor = RGB(192,192,192), ;
		Caption = "", ;
		Height = 16, ;
		Left = 18, ;
		Top = 27, ;
		Width = 319, ;
		WordWrap = .F., ;
		Name = "lblTask"


	ADD OBJECT shpthermbar AS shape WITH ;
		BorderStyle = 0, ;
		FillColor = RGB(128,128,128), ;
		FillStyle = 0, ;
		Height = 16, ;
		Left = 17, ;
		Top = 46, ;
		Width = 0, ;
		Name = "shpThermBar"


	ADD OBJECT lblpercentage AS label WITH ;
		FontName = WIN32FONT, ;
		FontSize = 8, ;
		BackStyle = 0, ;
		Caption = "0%", ;
		Height = 13, ;
		Left = 170, ;
		Top = 47, ;
		Width = 16, ;
		Name = "lblPercentage"


	ADD OBJECT lblpercentage2 AS label WITH ;
		FontName = WIN32FONT, ;
		FontSize = 8, ;
		BackColor = RGB(0,0,255), ;
		BackStyle = 0, ;
		Caption = "Label1", ;
		ForeColor = RGB(255,255,255), ;
		Height = 13, ;
		Left = 170, ;
		Top = 47, ;
		Width = 0, ;
		Name = "lblPercentage2"


	ADD OBJECT shape11 AS shape WITH ;
		BorderColor = RGB(128,128,128), ;
		Height = 1, ;
		Left = 16, ;
		Top = 45, ;
		Width = 322, ;
		Name = "Shape11"


	ADD OBJECT shape12 AS shape WITH ;
		BorderColor = RGB(255,255,255), ;
		Height = 1, ;
		Left = 16, ;
		Top = 61, ;
		Width = 323, ;
		Name = "Shape12"


	ADD OBJECT shape13 AS shape WITH ;
		BorderColor = RGB(128,128,128), ;
		Height = 16, ;
		Left = 16, ;
		Top = 45, ;
		Width = 1, ;
		Name = "Shape13"


	ADD OBJECT shape14 AS shape WITH ;
		BorderColor = RGB(255,255,255), ;
		Height = 17, ;
		Left = 338, ;
		Top = 45, ;
		Width = 1, ;
		Name = "Shape14"


	ADD OBJECT lblescapemessage AS label WITH ;
		FontBold = .F., ;
		FontName = WIN32FONT, ;
		FontSize = 8, ;
		Alignment = 2, ;
		BackStyle = 0, ;
		BackColor = RGB(192,192,192), ;
		Caption = "", ;
		Height = 14, ;
		Left = 17, ;
		Top = 68, ;
		Width = 322, ;
		WordWrap = .F., ;
		Name = "lblEscapeMessage"


*!*********************************************************************
*!
*!      Procedure: complete
*!
*!*********************************************************************
PROCEDURE complete
		* This is the default complete message
		parameters m.cTask
		if parameters() = 0
			m.cTask = THERMCOMPLETE_LOC
		endif
		this.Update(100,m.cTask)
ENDPROC


*!*********************************************************************
*!
*!      Procedure: update
*!
*!*********************************************************************
PROCEDURE update
		* m.iProgress is the percentage complete
		* m.cTask is displayed on the second line of the window

		parameters iProgress,cTask

		Local iPercentage, iAvgCharWidth
	
		if parameters() >= 2 .and. type('m.cTask') = 'C'
			* If we're specifically passed a null string, clear the current task,
			* otherwise leave it alone
			this.cCurrentTask = m.cTask
		endif
		
		if ! this.lblTask.Caption == this.cCurrentTask
			this.lblTask.Caption = this.cCurrentTask
		endif

		m.iPercentage = m.iProgress
		m.iPercentage = min(100,max(0,m.iPercentage))
		
		if m.iPercentage = this.iPercentage
			RETURN
		endif
		
		if len(alltrim(str(m.iPercentage,3)))<>len(alltrim(str(this.iPercentage,3)))
			iAvgCharWidth=fontmetric(6,this.lblPercentage.FontName, ;
				this.lblPercentage.FontSize, ;
				iif(this.lblPercentage.FontBold,'B','')+ ;
				iif(this.lblPercentage.FontItalic,'I',''))
			this.lblPercentage.Width=txtwidth(alltrim(str(m.iPercentage,3)) + '%', ;
				this.lblPercentage.FontName,this.lblPercentage.FontSize, ;
				iif(this.lblPercentage.FontBold,'B','')+ ;
				iif(this.lblPercentage.FontItalic,'I','')) * iAvgCharWidth
			this.lblPercentage.Left=int((this.shpThermBarMaxWidth- ;
				this.lblPercentage.Width) / 2)+this.shpThermBar.Left-1
			this.lblPercentage2.Left=this.lblPercentage.Left
		endif
		this.shpThermBar.Width = int((this.shpThermBarMaxWidth)*m.iPercentage/100)
		this.lblPercentage.Caption = alltrim(str(m.iPercentage,3)) + '%'
		this.lblPercentage2.Caption = this.lblPercentage.Caption
		if this.shpThermBar.Left + this.shpThermBar.Width -1 >= ;
			this.lblPercentage2.Left
			if this.shpThermBar.Left + this.shpThermBar.Width - 1 >= ;
				this.lblPercentage2.Left + this.lblPercentage.Width - 1
				this.lblPercentage2.Width = this.lblPercentage.Width
			else
				this.lblPercentage2.Width = ;
					this.shpThermBar.Left + this.shpThermBar.Width - ;
					this.lblPercentage2.Left - 1
			endif
		else
			this.lblPercentage2.Width = 0
		endif
		this.iPercentage = m.iPercentage
ENDPROC

*!*********************************************************************
*!
*!      Procedure: Init
*!
*!*********************************************************************
PROCEDURE Init
		* m.cTitle is displayed on the first line of the window
		* m.iInterval is the frequency used for updating the thermometer
		parameters cTitle, iInterval
		this.lblTitle.Caption = iif(empty(m.cTitle),'',m.cTitle)
		this.shpThermBar.FillColor = rgb(128,128,128)
		local cColor

		* Check to see if the fontmetrics for MS Sans Serif matches
		* those on the system developed. If not, switch to Arial. 
		* The RETURN value indicates whether the font was changed.
		if fontmetric(1, WIN32FONT, 8, '') <> 13 .or. ;
			fontmetric(4, WIN32FONT, 8, '') <> 2 .or. ;
			fontmetric(6, WIN32FONT, 8, '') <> 5 .or. ;
			fontmetric(7, WIN32FONT, 8, '') <> 11
			this.SetAll('FontName', WIN95FONT)
		endif

		m.cColor = rgbscheme(1, 2)
		m.cColor = 'rgb(' + substr(m.cColor, at(',', m.cColor, 3) + 1)
		this.BackColor = &cColor
		this.Shape5.FillColor = &cColor
ENDPROC

ENDDEFINE
