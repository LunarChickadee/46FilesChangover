___ PROCEDURE findname/5 _______________________________________________________
local firstname, lastname, clist, mlist
hide
noshow
firstrecord
loop
firstname=""
lastname=""
clist=0
mlist=0
firstname=Con[1," "][1,-2]
lastname=Con[" ",-1][2,-1]
clist=«C#»
window "37 mailing list"
find Con contains firstname and Con contains lastname
if info("found")=0
mlist=0
else
mlist=«C#»
endif
window "searchlist"
insertbelow
mailinglist=mlist
custhistory=clist
Con=firstname+" "+lastname
window "customer_history"
downrecord
until info("stopped")
show
___ ENDPROCEDURE findname/5 ____________________________________________________

___ PROCEDURE ccrider/ç ________________________________________________________
waswindow=info("windowname")
serverlookup "off"
NoUndo
GetScrap "enter the customer number"
Find «C#» = val(clipboard())
if info("found")=0
beep
endif
if info("files") contains "customer_history"
window "customer_history:secret"
Find «C#» = val(clipboard())
window waswindow
endif
;field «C#»
;field MAd
serverlookup "on"
___ ENDPROCEDURE ccrider/ç _____________________________________________________

___ PROCEDURE mailinglistlookup/1 ______________________________________________
custno=«C#»
window "41 mailing list"
find «C#»=custno
___ ENDPROCEDURE mailinglistlookup/1 ___________________________________________

___ PROCEDURE loopdown/2 _______________________________________________________
field «36Total»
Hide
loop
;copycell
;pastecell
downrecord
until 1000
Show
___ ENDPROCEDURE loopdown/2 ____________________________________________________

___ PROCEDURE fix money/3 ______________________________________________________
Field S31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field Bf31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field M31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field OGS31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
Field T31
loop
total
lastrecord
copy
find «C#»=cusno1
paste
right
stoploopif val(info("fieldname")[-2,-1])=val(str(«C#»)[1,2])-1
until forever
RemoveSummaries 7
find «C#»≠cusno1
cusno2=«C#»
deleterecord
window "33 mailing list"
find «C#»=cusno2
deleterecord
___ ENDPROCEDURE fix money/3 ___________________________________________________

___ PROCEDURE findorder/4 ______________________________________________________
getscrap "What order? (Please use all 5 digits)"
openfile "37orders"
ono=val(clipboard())
find OrderNo=ono
case (ono≥10000 and ono<30000) or (ono>60000 and ono<100000) or (ono>800000 and ono<1000000)
GoForm "seedsinput"
case (ono>30000 and ono<40000) or (ono>300000 and ono<400000)
goform "ogsinput"
case ono>40000 and ono<50000
goform "treesinput"
case ono>50000 and ono<60000
goform "bulbsinput"
case ono>70000 and ono<80000
goform "mtinput"
endcase

___ ENDPROCEDURE findorder/4 ___________________________________________________

___ PROCEDURE forceunlock ______________________________________________________
forceunlockrecord
___ ENDPROCEDURE forceunlock ___________________________________________________

___ PROCEDURE .Initialize ______________________________________________________
global custno
custno=0
;GoSheet
;Field "MAd"
;SortUp
;Field "City"
;SortUp
;Field "St"
;SortUp
;Field "Zip"
;SortUp
;windowtoback "customer_history"


___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE filladdress ______________________________________________________
select MAd=""
if info("selected")=info("records")
message "all uptodate"
stop
endif
field MAd
formulafill lookup("38 mailing list","C#",«C#»,"MAd","",0)
field City
formulafill lookup("38 mailing list","C#",«C#»,"City","",0)
field St
formulafill lookup("38 mailing list","C#",«C#»,"St","",0)
field Zip
formulafill lookup("38 mailing list","C#",«C#»,"Zip",0,0)
___ ENDPROCEDURE filladdress ___________________________________________________

___ PROCEDURE .find ____________________________________________________________
custno=«C#»
window "37 mailing list"
find «C#»=custno
___ ENDPROCEDURE .find _________________________________________________________

___ PROCEDURE sortup ___________________________________________________________
field MAd
sortup
field City
sortup
field St
sortup
field Zip
sortup
___ ENDPROCEDURE sortup ________________________________________________________

___ PROCEDURE consolidate ______________________________________________________
local mlist, clist
global dialogPause
window "searchlist"
loop
mlist=mailinglist
clist=custhistory
window "customer_history:customer history"
select «C#»=mlist
selectadditional «C#»=clist
find «C#»=clist
cancelok "Delete This record"
if clipboard()="OK"
deleterecord
endif
window "searchlist"
downrecord
until info("stopped")
___ ENDPROCEDURE consolidate ___________________________________________________

___ PROCEDURE delete ___________________________________________________________
lastrecord
loop
deleterecord
until info("selected")=1
field «C#»
copy
selectall
find «C#»=clipboard()
___ ENDPROCEDURE delete ________________________________________________________

___ PROCEDURE DeleteRecord _____________________________________________________
deleterecord
window "searchlist"
downrecord
call "consolidate"

___ ENDPROCEDURE DeleteRecord __________________________________________________

___ PROCEDURE check address ____________________________________________________
field MAd
select MAd notmatch  lookup("37 mailing list","C#",«C#»,"MAd","",0)
___ ENDPROCEDURE check address _________________________________________________

___ PROCEDURE fill info ________________________________________________________
forcesynchronize
window "45 mailing list"
call "forcesynchronize"
window "customer_history"
select Zip=0 and length(St)=2
if info("selected")=info("records")
beep
stop
endif
field Zip
formulafill lookup("45 mailing list", "C#",«C#», "Zip",0,0)
select MAd=""
field MAd
formulafill lookup("45 mailing list", "C#",«C#», "MAd","",0)
select City=""
field City
formulafill lookup("45 mailing list", "C#",«C#», "City","",0)
select St=""
field St
formulafill lookup("45 mailing list", "C#",«C#», "St","",0)
selectall
call "sortup"
select MAd=""
___ ENDPROCEDURE fill info _____________________________________________________

___ PROCEDURE forcesynchronize _________________________________________________
forcesynchronize
call "sortup"
___ ENDPROCEDURE forcesynchronize ______________________________________________

___ PROCEDURE checkit __________________________________________________________
local checktotal
checktotal=0
firstrecord
loop
checktotal=checktotal+val(«Gets Check»)
downrecord
until info("stopped")
message str(checktotal)
___ ENDPROCEDURE checkit _______________________________________________________

___ PROCEDURE fixaddress _______________________________________________________
field MAd
select MAd≠lookup("37 mailing list", "C#",«C#», "MAd","",0)
formulafill lookup("37 mailing list", "C#",«C#», "MAd","",0)
field City
formulafill lookup("37 mailing list", "C#",«C#», "City","",0)
field St
formulafill lookup("37 mailing list", "C#",«C#», "St","",0)
field Zip
formulafill lookup("37 mailing list", "C#",«C#», "Zip",0,0)
selectall
call "sortup"
___ ENDPROCEDURE fixaddress ____________________________________________________

___ PROCEDURE fix zipcode ______________________________________________________
select Zip=0 and length(St)=2
field Zip
formulafill lookup("37 mailing list","C#",«C#»,"Zip",0,0)
___ ENDPROCEDURE fix zipcode ___________________________________________________

___ PROCEDURE selectcustomers __________________________________________________
getscrap "Which division"
case clipboard() contains "bulbs"
select «Bf33» >0 or «Bf34» >0 or «Bf35» >0 or «Bf36» >0 
selectwithin «C#»>0
case clipboard() contains "trees"
select «T33» >0 or «T34» >0 or «T35» >0  or «T36» >0
selectwithin «C#»>0
case clipboard() contains "seeds"
select «S33» >0 or «S34» >0 or «S35» >0 or «S36» >0
selectadditional «MT33» >0 or «MT34» >0 or «MT35» >0 or «MT36» >0
selectadditional «OGS33» >0 or «OGS34» >0 or «OGS35» >0 or «OGS36» >0
selectwithin «C#»>0
endcase
window "37 mailing list"
field «C#»
select «C#»=lookupselected("customer_history","C#",«C#»,"C#",0,0) and «C#»>0
___ ENDPROCEDURE selectcustomers _______________________________________________

___ PROCEDURE close window _____________________________________________________
SelectAll
save
CloseFile
___ ENDPROCEDURE close window __________________________________________________

___ PROCEDURE huh ______________________________________________________________
field T43
noshow

formulafill lookupselected("43treestally-DL 6-29-21","C#",«C#»,"AdjTotal",0,0)
endnoshow
___ ENDPROCEDURE huh ___________________________________________________________

___ PROCEDURE .ExportMacros ____________________________________________________


local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros are saved to your clipboard!"
___ ENDPROCEDURE .ExportMacros _________________________________________________
