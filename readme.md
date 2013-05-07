ParseLocals.prg

Installation:
Do install just run the ParseLocals.prg in your Visual FoxPro startup program.
It will install itself at the bottom of the Tools menu. 

Usage:
Open the source code you want to check in the editor.
If you have been editing there is no need save. ParseLocals copies the code from the editor not the file from disk.
Select "Check local declarations" from the Tools menu or hit the hotkey of Alt+L.
If ParseLocals detects any omissions in the local declarations it will report them per method or procedure.
	The section following Add Local indicates variables that where assigned a value but for which no declaration was found
	The section following Remove Local indicates variables for which a declaration was found but that appear to be unused.
	If no omissions are detected ParseLocals is silent.

If you select any code in the editor before hitting Alt+L only the selected code will be parsed. You can use this feature to check a single function or method in a large code base class or procedure file.
	
The parser does its best to analyse the code in question and report on the variables as best as possible. There is however no guarantee that all omissions are detected and all detected omissions are correct. Please use common sense and keep in mind that ParseLocals isn't 100% reliable just 99%.

The parser is used at design time so has no knowledge of the run time environment. For that reason it cannot take into account variables declared as public or private at a higher level and will report these as missing local declarations. To remove these incorrect messages you can declare these using "*:local" (without the quotes). Visual FoxPro will ignore this as a comment but ParseLocal will treat it as a normal local declaration.

The ParseLocal parser serves the same purpose as using the _VFP.LanguageOptions = 1 setting in detecting variables that aren't declared. The main difference between the two is the approach taken. LanguageOptions is a runtime check where you must run the code to check it and each undeclared variable generates a warning message in the debug output window. ParseLocals is a design time check where the code is checked while the developer is editing. Each approach has pros and cons:

LanguageOptions
Pros:
-	It knows about private and public variables and macro execution so all the errors reported are guaranteed to be errors.
Cons:
-	It only checks code as it is executing so only code that is really executed is checked. Code that is passed because of failed conditions is not checked so it is hard to do a complete check of all variables used. 
-	The developer must run the code, check for errors, edit the code and restart the application. This makes for a much slower test cycle.
-	There is no check for variables that are declared but never used.

ParseLocals
Pros:
-	All code is checked even if it is never executed. This helps finding typos in error handling code that is normally not executed.
-	The developer only needs to hit the hot key in the editor to check the current code and see all errors. This makes for a much faster test and fix cycle.
-	Variables that are declared nut never used are reported as well.
Cons:
-	The parser isn't aware of private and public variables available at runtime.
-	The parser isn't able to parse macro executed code so variables created in a macro will not be reported.


Note: The parser assumes the code is valid Visual FoxPro source code. If the code isn't valid errors may occur.


Any feedback will be appreciated, email me at mauricedb@computer.org 
Future updates will be posted on my web site: http://www.TheProblemSolver.nl

Copyright.
All source code and documentation contained in this zip file has been 
placed into the public domain. All files contained in this zip file 
are provided 'as is' without warranty of any kind.  In no event shall 
its authors, contributors, or distributors be liable for any damages.
