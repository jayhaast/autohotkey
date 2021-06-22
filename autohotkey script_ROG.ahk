; GLOBAL COMMANDS NEEDED TO LOAD FIRST::

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode 2   ;sets window title matching (for #IfWinActive) to partial matches

; add file explorer to a group called 'Explorer', as it has multiple ahk classes
GroupAdd, Explorer, ahk_class CabinetWClass
GroupAdd, Explorer, ahk_class ExploreWClass
Return
;=====================================================================================================
;=====================================================================================================

;-------------------------
; F KEYS
;-------------------------

f1:: 
  Send {LWin}
  Sleep 100
  Send Focus
  Sleep 100
  Send {enter}
Return

;f1:: Run C:\dev\r
f2:: Run C:\Users\User\OneDrive - Haast Energy Trading\Documents\onboarding docs\windfarm model\ ;temporarily point to main folder i'm working in
f3:: Run C:\Users\User\OneDrive - Haast Energy Trading\Documents\onboarding docs
;f9:: Send !{tab}
f5:: Send !{f4}				
F7:: Run C:\Users\User\AppData\Local\Microsoft\WindowsApps\Slack.exe

f8:: ; winkey, because not working on keyboard. Plus extra code to stop Dragon popping up
 Send {LWin}
 Sleep 500
 Send a
 Sleep 500
 Send {backspace}
Return

f9:: 
 Run C:\Windows\system32\SnippingTool.exe
 Sleep 1000
 WinActivate "Snipping Tool"
 Sleep 1000
 Send ^n
Return

; CHROME: switch to open program if there is one, otherwise open
f12:: 
if WinExist("Chrome")
  WinActivate 
else
  Run C:\Program Files\Google\Chrome\Application\chrome.exe
Return

; RSTUDIO: switch to open program if there is one, otherwise open
^f1:: 
if WinExist("RStudio")
  WinActivate 
else
  Run C:\Program Files\RStudio\bin\rstudio.exe
Return

; NOTEPAD: switch to open program if there is one, otherwise open
^f2:: 
if WinExist("Notepad")
  WinActivate 
else
  Run C:\Windows\system32\notepad.exe
Return

; WORD: switch to open program if there is one, otherwise open
^f3:: 
if WinExist("Word")
  WinActivate 
else
  Run "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
Return

; FILE EXPLORER: switch to open program if there is one, otherwise open
^f4:: 
if WinExist("ahk_group Explorer") ;group is defined at top of script
  WinActivate 
else
  Send #{e}
Return

;-------------------------
; HOTKEYS
;-------------------------

:*:`;tj::{enter}{enter}Thanks,{enter}{enter}Jay

:*:`;cf:: ;confluence. "general" and "r" spaces only are relevant to me
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://haastenergy.atlassian.net/wiki/spaces
Return

:*:`;tm:: ;time management via jira
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://haastenergy.atlassian.net/secure/RapidBoard.jspa?rapidView=5&projectKey=GEN&view=planning&selectedIssue=GEN-473&issueLimit=100
Return

:*:`;gm:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://gmail.com
Return

:*:`;gc:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://calendar.google.com/calendar/u/0/r/week
Return

:*:`;do:: 
 Run C:\Users\User\Documents
Return

:*:`;dl:: 
 Run C:\Users\User\Downloads
Return

:*:`;ri:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://ap-southeast-2.console.aws.amazon.com/ec2/v2/home?region=ap-southeast-2#Instances:sort=instanceState ;ec2 running instances
Return

:*:`;s3:: 
 Run C:\Program Files\Google\Chrome\Application\chrome.exe https://s3.console.aws.amazon.com/s3/buckets/nz-omg-ann-aipl-staging/digital_pathway_reporting/?region=ap-southeast-2&tab=overview ;S3
Return

; Opening via hotstrings - other

:*:`;jh::jay@haastenergy.com
:*:`;rj::ruffell.jay@gmail.com
:*:`;fp::Spring@21{!}
:*:`;sp::altheg81
:*:`;021::0210528720

;#left::
;WinGet, mm, MinMax, A
;WinRestore, A
;WinGetPos, X, Y,,,A
;WinMove, A,, X-A_ScreenWidth, Y
;if(mm = 1){
;    WinMaximize, A
;}
;return
;					;moves window from right to left monitor
;#right::
;WinGet, mm, MinMax, A
;WinRestore, A
;WinGetPos, X, Y,,,A
;WinMove, A,, X+A_ScreenWidth, Y
;if(mm = 1){
;    WinMaximize, A
;}
;return
					
;-------------------------
; MODIFIER KEYS
;-------------------------

; Because Dragon is a pain with the Windows key, add a quick letter and backspace to stop Dragon opening
LWin::
 Send {LWin}
 Sleep 500
 Send a
 Sleep 500
 Send {backspace}
Return

; remap winkey + xxx keys to right control (keyboard winkey not working)  
>^up::  
 Send #{up}
Return

>^down::
 Send #{down}
Return

>^right::
 Send #{right}
Return

>^left::
 Send #{left}
Return					

;=====================================================================================================
;=====================================================================================================

;RStudio

#IfWinActive RStudio			; applies only to RStudio from now on

;-------------------------
; Dplyr commands
;-------------------------
:*:`;nl::
:*:`;nl::             ; 'next line' - when cursor is in brackets in a dplyr pipe, go to next line
 Send {end}
 Send {enter}
Return

::sel::               ; select() %>% with cursor finishing in brackets
 Send select(){space}
 Send +5
 Send +.
 Send +5
 Send {left 5}
Return
::gb::               ; select() %>% with cursor finishing in brackets
 Send group_by(){space}
 Send +5
 Send +.
 Send +5
 Send {left 5}
Return
::mut::               ; select() %>% with cursor finishing in brackets
 Send mutate(){space}
 Send +5
 Send +.
 Send +5
 Send {left 5}
Return
::tra::               ; select() %>% with cursor finishing in brackets
 Send transmute(){space}
 Send +5
 Send +.
 Send +5
 Send {left 5}
Return
::sz::               ; select() %>% with cursor finishing in brackets
 Send summarise(){space}
 Send +5
 Send +.
 Send +5
 Send {left 5}
Return
::fil::               ; select() %>% with cursor finishing in brackets
 Send filter(){space}
 Send +5
 Send +.
 Send +5
 Send {left 5}
Return

::oo::               ; pipeline operator
 Send +5
 Send +.
 Send +5
 Send {enter}
Return

::ooo::               ; as above but no enter at end - useful when inserting a line within an existing pipe.
 Send +5
 Send +.
 Send +5
 Send {down}
 Send {home}
Return

:*:`;dd:: ; dd %>%
 Send dd{space}
 Send +5
 Send +.
 Send +5
 Send {enter}
Return

;-------------------------
; Other Rstudio commands
;-------------------------

^+a:: ; change appearance
 Send !t
 Sleep 200
 Send g
 Sleep 400
 Send {down 3}
; Send {enter}
Return

^w::
 Send ^{w}				;close tab
Return

RAlt:: 					; send lines with right alt key
  Send ^{enter}
Return

^tab:: 				;switch tabs
  Send ^!{down}
Return

^e::				;cancel operation (switch to code window, press escape, and switch back to source)
 Send ^2
 Send {escape}
 Send ^1
Return

^n:: Send {Home}+{End}			;highlight line:

;Mouseclick scripts:

 ~LButton::
 If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500)
 {
  Send ^{left} 									;double click selects word
  Send ^+{right}
 }
 Return

 ~RButton::
 If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500)
 {
  Sleep 50 ; wait for right-click menu, fine tune for your PC		;double right click does 'paste click'
  Send {Esc} ; close it
  Send ^v ; or your double-right-click action here
 }
 Return


;HOTSTRINGS:

::`;t:: 
 Send traceback() 
 Send {enter}
Return

:*:`;hs::
 Send ^+{home}
 Sleep 200 
 Send ^{enter}
 Sleep 500
 Send {down}
Return					;'highlight start' [but actually 'buy from start']

::`;s::               ;'str()
 Send str()
 Send {left}
Return

::`;h::               ;'head('
 Send head()
 Send {left}
Return

::hh::               ;hashkey followed by space.
 Send +3
 Send {space}
Return

:*:`;fu::             ;'full  underline'
 Send +3
 Send __________________________________________________________
 Send {enter}
 Send {enter}
Return

:*:`;el::             ;'equal line'
 Send +3
 Send {+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}{+}
 Send {enter}
Return

::`;t::
 Send tempDF                ;note you have to push enter or space for the tempDFs... cost otherwise you can't do 'tempDF$' without first activating 'tempDF'
Return

::`;td::
 Send tempDF$
Return
 
:*:,,::`<- `		;hit comma twice to assign variable

;#### temporary bookmarking:

:*:`;mh::				; "Mark Here". "*" means no ending character required to activate
 Send {home}				
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
Return					

:*:`;gm::				; "Go Mark". "*" means no ending character required to activate
 Send ^f				
 Sleep 200
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
 Sleep 200
 Send {escape}
Return		


;# BOOKMARK NAVIGATION=====================================

::`;b1::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 1.	
 Sleep 200
 Send {escape}
Return		


::`;b2::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 2.
 Sleep 200
 Send {escape}
Return		


::`;b3::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 3.	
 Sleep 200
 Send {escape}
Return		

::`;b4::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 4.
 Sleep 200
 Send {escape}
Return		

::`;b5::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 5.
 Sleep 200
 Send {escape}
Return		

::`;b6::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 6.
 Sleep 200
 Send {escape}
Return		

::`;b7::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 7.	
 Sleep 200
 Send {escape}
Return		


::`;b8::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 8.
 Sleep 200
 Send {escape}
Return		


::`;b9::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 9.	
 Sleep 200
 Send {escape}
Return		

::`;b10::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 10.
 Sleep 200
 Send {escape}
Return		

::`;b11::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 11.
 Sleep 200
 Send {escape}
Return		

::`;b12::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 12.
 Sleep 200
 Send {escape}
Return		

::`;b13::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 13.
 Sleep 200
 Send {escape}
Return		

::`;b14::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 14.
 Sleep 200
 Send {escape}
Return		

::`;b15::				;press comma, b, m, <number> to navigate to bookmark
 Send ^f
 Sleep 200
 Send ^a {del}			
 Send +3 			;hash
 Sleep 200 
 Send {home} 
 Send {delete} 
 Send {end}	;deletes space automatically insterted before hash
 Send 15.
 Sleep 200
 Send {escape}
Return		


#IfWinActive
					
;=====================================================================================================
;=====================================================================================================

;WORD

#IfWinActive ahk_class OpusApp ; applies only to Word from now on

^!p::
Send ^a
Send ^x
WinActivate,RStudio
Send ^v
Return				; cut and paste to TinnR/RStudio 



;#### temporary bookmarking:

:*:`;mh::				; "Mark Here". "*" means no ending character required to activate
 Send {home}				
 Send `===== TEMPORARY BOOKMARK =====	
Return					

:*:`;gm::				; "Go Mark". "*" means no ending character required to activate
 Send ^f				
 Sleep 200
 Send `===== TEMPORARY BOOKMARK =====	
 Send {enter}
 Sleep 200
 Send {escape}
Return		

#IfWinActive

;=====================================================================================================
;===================================================================================================

;WINDOWS EXPLORER

#IfWinActive ahk_class ExploreWClass 
#w::               
 Send !d
Return                                             ; highlights address bar. Note hotstrings don't work here.

#IfWinActive

#IfWinActive ahk_class CabinetWClass               ;2 possible window names

#w::               
 Send !d
Return

#IfWinActive


;=====================================================================================================
;=====================================================================================================
;===================================================================================================

;OUTLOOK

#IfWinActive Outlook 

:*:`;hg::
 Send Hey guys,
 Send {enter}{enter}			
Return

:*:`;hc::
 Send Hey Charles,
 Send {enter}{enter}			
Return

:*:`;tj::
 Send {enter}{enter}
 Send Thanks,
 Send {enter}{enter}
 Send Jay				
Return

#IfWinActive


;=====================================================================================================;=====================================================================================================
;===================================================================================================

;OUTLOOK - new message window

#IfWinActive Message 

:*:`;hg::
 Send Hey guys,
 Send {enter}{enter}			
Return

:*:`;hc::
 Send Hey Charles,
 Send {enter}{enter}			
Return

:*:`;tj::
 Send {enter}{enter}
 Send Thanks,
 Send {enter}{enter}
 Send Jay				
Return

#IfWinActive

;=====================================================================================================
;=====================================================================================================
					
;Spyder

#IfWinActive ahk_exe pythonw.exe
::hh::               ;hashkey followed by space.
 Send +3
 Send {space}
Return

#IfWinActive

;=====================================================================================================
;=====================================================================================================

;=====================================================================================================
;=====================================================================================================
;=====================================================================================================
;=====================================================================================================
