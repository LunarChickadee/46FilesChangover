___ PROCEDURE .Initialize ______________________________________________________
field Date
sortup
field OrderNo
sortupwithin
___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE hidewindow/h _____________________________________________________
Window "Hide This Window"
window waswindow
___ ENDPROCEDURE hidewindow/h __________________________________________________

___ PROCEDURE export ___________________________________________________________
export "shipping.txt", "'"+Contact+"'"+","+"'"+Group+"'"+","+"'"+«Sh Add1»+"'"+","+"'"+«Sh Add2»+"'"+","+"'"+City+"'"+","+"'"+St+"'"+","
    +"'"+pattern(Zip,"#####")+"'"+","+str(«O#»)+","+"'"+«UPS#»+"'"+","+«C#»+","+"'"+Code+"'"+","+"'"
    +datepattern(Date, "mm/dd/yy")+"'"+","+str(Weight)+","+""+"\r"+¶
selectall
lastrecord
___ ENDPROCEDURE export ________________________________________________________

___ PROCEDURE .closewindow _____________________________________________________
CloseWindow
___ ENDPROCEDURE .closewindow __________________________________________________

___ PROCEDURE just_locate/4 ____________________________________________________
global Numb
getscrap "What's the order number?"
Numb=int(val(striptonum(clipboard())))
find OrderNo=Numb
___ ENDPROCEDURE just_locate/4 _________________________________________________

___ PROCEDURE huh ______________________________________________________________
message val(«UPS#»[-4,-1])
___ ENDPROCEDURE huh ___________________________________________________________

___ PROCEDURE track/6 __________________________________________________________
local ono
gettext "Which order?",ono
find OrderNo=val(ono)
clipboard()=""
if «UPS#» beginswith "1Z"
clipboard()="https://wwwapps.ups.com/WebTracking/track?loc=en_US&track.x=Track&trackNums="+«UPS#»
else
if «UPS#» beginswith "16" or («UPS#» beginswith "6" and «UPS#» notcontains "ross")
clipboard()="https://www.fedex.com/apps/fedextrack/index.html?tracknumbers="+«UPS#»
else
    if «UPS#» endswith "US" OR (length(«UPS#»[-4,-1])=4 and val(«UPS#»[-4,-1])>0)
clipboard()="https://tools.usps.com/go/TrackConfirmAction?tLabels="+«UPS#»
    endif
    endif
endif
    if clipboard()≠""
applescript |||
tell application "Finder"
	activate
	open location the clipboard
end tell
|||
    endif

___ ENDPROCEDURE track/6 _______________________________________________________

___ PROCEDURE tallylookup-trees ________________________________________________
global Numb
Numb=int(OrderNo)
openfile "45treestally"
selectall
resynchronize
select OrderNo=Numb
if info("selected")<info("records")
show
openform "treespagecheck"
stop
else
message "nothing found"
endif
___ ENDPROCEDURE tallylookup-trees _____________________________________________

___ PROCEDURE tallylookup-OGS __________________________________________________
global Numb
Numb=int(OrderNo)
openfile "45ogstally"
selectall
resynchronize
select OrderNo=Numb
if info("selected")<info("records")
show
    if OrderNo<60000
    openform "ogspagecheck"
    else
    openform "mtpagecheck"
    endif
stop
else
message "nothing found"
endif
___ ENDPROCEDURE tallylookup-OGS _______________________________________________

___ PROCEDURE tallylookup-bulbs ________________________________________________
global Numb
Numb=int(OrderNo)
openfile " 45bulbstally"
selectall
resynchronize
select OrderNo=Numb
if info("selected")<info("records")
show
openform "bulbspagecheck"
stop
else
message "nothing found"
endif
___ ENDPROCEDURE tallylookup-bulbs _____________________________________________

___ PROCEDURE tallylookup-seeds ________________________________________________
global Numb
Numb=int(OrderNo)
openfile "45seedstally"
selectall
resynchronize
select OrderNo=Numb
if info("selected")<info("records")
show
openform "seedspagecheck"
stop
else
message "nothing found"
endif
___ ENDPROCEDURE tallylookup-seeds _____________________________________________

___ PROCEDURE track ____________________________________________________________
clipboard()=""
if «UPS#» beginswith "1Z"
clipboard()="https://wwwapps.ups.com/WebTracking/track?loc=en_US&track.x=Track&trackNums="+«UPS#»
else
if «UPS#» beginswith "16" or («UPS#» beginswith "6" and «UPS#» notcontains "ross")
clipboard()="https://www.fedex.com/apps/fedextrack/index.html?tracknumbers="+«UPS#»
else
    if «UPS#» endswith "US" OR (length(«UPS#»[-4,-1])=4 and val(«UPS#»[-4,-1])>0)
clipboard()="https://tools.usps.com/go/TrackConfirmAction?tLabels="+«UPS#»
    endif
    endif
endif
    if clipboard()≠""
applescript |||
tell application "Finder"
	activate
	open location the clipboard
end tell
|||
    endif

___ ENDPROCEDURE track _________________________________________________________

___ PROCEDURE .ExportMacros ____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros from " +info("databasename")+" are saved to your clipboard!"
___ ENDPROCEDURE .ExportMacros _________________________________________________
