___ PROCEDURE transferCity/† ___________________________________________________
local address
arraylinebuild address,¬, info("DatabaseName") , City+¬+St+¬+str(Zip)+¬+str(adc)+¶

openfile "Inquiry List"
GoForm "inquiryentry"
InsertBelow
Field «Con»

City=extract(address, ¬,1)
St=extract(address, ¬,2)
Zip=val(extract(address, ¬,3))
adc=val(extract(address, ¬,4))

___ ENDPROCEDURE transferCity/† ________________________________________________

___ PROCEDURE AlreadyOnMove2inqList/å __________________________________________
orders=""
ArrayLineBuild orders,¶, info("DatabaseName"), " "+¬+str(«C#»)+¬+Con+¬+Group+¬+MAd+¬+City+¬+St+¬+str(Zip)+¬+str(adc)
openfile "Inquiry List"
openfile "+@orders"
GoForm "entrybox"
field «M?» 
editcell

___ ENDPROCEDURE AlreadyOnMove2inqList/å _______________________________________

___ PROCEDURE .Initialize ______________________________________________________

Global New, Num, Numb, enter, place, ID, CC, ED, credit, orders, ship, Flag, vFacWindow, ono

Num=0
zoomWindow 23,0,250,1496,""
NoUndo
vFacWindow = info("windowname")
NoShow
openFile "45 mailing list"
    SelectAll
    
    call "listsortcomplete/1"
    save
    //CloseFile
    
window vFacWindow

OpenFile "&&45 mailing list"
OpenForm "view orders blue box"
   zoomwindow 285,0,494,802,""
   OpenFile "&&45 mailing list"

save
EndNoShow

OpenSecret "newadc"
    
;Openfile "customer_history"
;    zoomwindow 783,0,410,647,""
;   save
;    CloseFile

Openfile "45orders"
    goform "facilitator view"
    zoomwindow 25,1159,764,760,""
    save



   ; OpenForm "functions"
  ; zoomwindow 885,850,153,241,""
  
Openfile "45shipping"
    goform "packages"
    zoomwindow 868,1223,326,701,""
    save
    
Openfile "45seedstally-winter"
   zoomwindow 83,959,951,932,""
    field OrderNo
    sortup
    save

Openfile "45ogstally"

    // window "45ogstally:bulbspagecheck"
   // zoomwindow 163,903,899,870,""
    field OrderNo
    sortup
    save

/*
Openfile "45treestally"
   zoomwindow 118,931,822,862,""
    field OrderNo
    sortup
    save
*/
;Openfile "45bulbstally"
;   zoomwindow 198,961,822,862,""
;    field OrderNo
;    sortup
;    save
openfile "ycitylist"
;window  "45 facilitator mailing list:functions"
window  "45 facilitator mailing list:view orders blue box"
message "all done opening!"
___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE .addnames ________________________________________________________
openfile "Inquiry List"


OpenForm "inquiryentryView"
    zoomwindow 300,850,274,852,""
OpenForm "entrybox"
    zoomwindow 583,850,400,800,""


___ ENDPROCEDURE .addnames _____________________________________________________

___ PROCEDURE .sendemail _______________________________________________________
local mailarray
mailarray=""
case mailcopies≠""
mailarray=mailaddress+","+mailcopies
sendarrayemail "", mailarray, mailheader, messageBody
case mailcopies=""
sendoneemail "", mailaddress, mailheader, messageBody
endcase
___ ENDPROCEDURE .sendemail ____________________________________________________

___ PROCEDURE updateEdit _______________________________________________________
if «C#»=0
    message "Sorry, cannot update non-customer record.  Put aside."
    stop
    endif

Num=«C#»
waswindow = info("windowname")

OpenFile "45 mailing list"
SelectAll
Select «C#»=Num
    if info("found")=0
    message "Sorry, something's wrong with this customer number. Please look deeper or get help."
    stop
    endif
SetWindow 350,350,410, 500, "noHorzScroll noVertScroll"
OpenForm "addresschecker"
YesNo "This is the one, yes?"
    if clipboard() contains "No"
    stop
    endif
CloseWindow
GoSheet    

Con=grabdata("45 facilitator mailing list", Con)
Group=grabdata("45 facilitator mailing list", Group)
MAd=grabdata("45 facilitator mailing list", MAd)
City=grabdata("45 facilitator mailing list", City)
St=grabdata("45 facilitator mailing list", St)
Zip=grabdata("45 facilitator mailing list", Zip)
adc=grabdata("45 facilitator mailing list", adc)
«M?»=grabdata("45 facilitator mailing list", «M?»)
SAd=grabdata("45 facilitator mailing list", SAd)
Cit=grabdata("45 facilitator mailing list", Cit)
Sta=grabdata("45 facilitator mailing list", Sta)
Z=grabdata("45 facilitator mailing list", Z)
phone=grabdata("45 facilitator mailing list", phone)
email=grabdata("45 facilitator mailing list", email)
S=grabdata("45 facilitator mailing list", S)
Bf=grabdata("45 facilitator mailing list", Bf)
T=grabdata("45 facilitator mailing list", T)
RedFlag=grabdata("45 facilitator mailing list", RedFlag)
Notes=grabdata("45 facilitator mailing list", Notes)
;save ;knocked out 2-3-15 gene
SelectAll

window "Hide This Window"

window "45 facilitator mailing list:Edit Record"
CloseWindow
window "45 facilitator mailing list:view orders blue box"

___ ENDPROCEDURE updateEdit ____________________________________________________

___ PROCEDURE .customerinfo ____________________________________________________
global cname, cgroup, cmad, ccity, czip
if «C#»=0
    AlertOk "Sorry, you cannot edit a non-customer record. Put info aside for now."
    stop
    endif

setwindowrectangle rectanglesize(681,136,320,842),""
OpenForm "Edit Record"
___ ENDPROCEDURE .customerinfo _________________________________________________

___ PROCEDURE (.tab ____________________________________________________________
;If info("trigger")="Key.Return" or info("trigger")="Button.CloseWindow"
;YesNo "Do yo want to update the real mailing list?"
;If clipboard()="Yes"
;if «C#»>0
;Num=«C#»
;else place=Con
;endif
;window "35 mailing list:Secret"
;openSecret "35 mailing list"
;if Num>0
;find Num=«C#»
;else 
;find Con=place
;    if info("found")=0
;    YesNo "Open Inquiry List"
;        if clipboard()="Yes"
;        openfile "Inquiry Lis"
;        insertbelow
;        field Con
;        editcellstop
;        endif
;    endif
;endif    
;Con=grabdata("35 facilitator mailing list", Con)
;Group=grabdata("35 facilitator mailing list", Group)
;MAd=grabdata("35 facilitator mailing list", MAd)
;City=grabdata("35 facilitator mailing list", City)
;St=grabdata("35 facilitator mailing list", St)
;Zip=grabdata("35 facilitator mailing list", Zip)
;adc=grabdata("35 facilitator mailing list", adc)
;phone=grabdata("35 facilitator mailing list", phone)
;email=grabdata("35 facilitator mailing list", email)
;window "Hide This Window"
;EndIf
;Endif
;window waswindow
___ ENDPROCEDURE (.tab _________________________________________________________

___ PROCEDURE (.findcustomer ___________________________________________________
; does this get used ??

if info("trigger")="Button.Find Customer"
if info("files") contains "37orders"
goto continue
else
opensecret "37orders"
endif
continue:
endif
loop
getscrap "Name contains"
find Con contains clipboard()
if info("found")=0
beep
endif
repeatloopif info("found")=0
stoploopif info("found")
while forever
Stop

___ ENDPROCEDURE (.findcustomer ________________________________________________

___ PROCEDURE .findorder _______________________________________________________
global forder
forder=0
forder=val(extract(officeorders,"  ",1))
window "45orders:facilitator view"
find OrderNo=forder
___ ENDPROCEDURE .findorder ____________________________________________________

___ PROCEDURE (.find35order ____________________________________________________
global f35order
forder=0
forder=val(extract(35officeorders,"  ",1))
window "35officeorders:facilitator view"
find OrderNo=f35order
___ ENDPROCEDURE (.find35order _________________________________________________

___ PROCEDURE .findseeds _______________________________________________________
global fseeds
fseeds=0
fseeds=val(extract(seedorders,"  ",1))
window " 45seedstally:facilitator view"
find OrderNo=fseeds
___ ENDPROCEDURE .findseeds ____________________________________________________

___ PROCEDURE (.find37seeds ____________________________________________________
global fseeds
fseeds=0
fseeds=val(extract(seedorders37,"  ",1))
window "37seedstally_delinked:facilitator view"
find OrderNo=fseeds
___ ENDPROCEDURE (.find37seeds _________________________________________________

___ PROCEDURE .findogs _________________________________________________________
global fogs
fogs=0
fogs=val(extract(ogsorders,"  ",1))
window "45ogstally:facilitator view"
select OrderNo=fogs
___ ENDPROCEDURE .findogs ______________________________________________________

___ PROCEDURE (.find37ogs ______________________________________________________
global fogs
fogs=0
fogs=val(extract(ogsorders37,"  ",1))
window "37ogstally_delinked:facilitator view"
select OrderNo=fogs
___ ENDPROCEDURE (.find37ogs ___________________________________________________

___ PROCEDURE .findtrees _______________________________________________________
global ftrees
ftrees=0
ftrees=val(extract(treeorders,"  ",1))
window "45treestally:facilitator view"
find OrderNo=ftrees
___ ENDPROCEDURE .findtrees ____________________________________________________

___ PROCEDURE (.findtrees37 ____________________________________________________
global ftrees
ftrees=0
ftrees=val(extract(treeorders37,"  ",1))
window "37treestally_delinked:facilitator view"
find OrderNo=ftrees
___ ENDPROCEDURE (.findtrees37 _________________________________________________

___ PROCEDURE .findbulbs _______________________________________________________
global fbulbs
fbulbs=0
fbulbs=val(extract(bulborders,"  ",1))
window "45bulbstally:facilitator view"
find OrderNo=fbulbs
___ ENDPROCEDURE .findbulbs ____________________________________________________

___ PROCEDURE (.findbulbs37 ____________________________________________________
global fbulbs
fbulbs=0
fbulbs=val(extract(bulborders37,"  ",1))
window "37bulbstally_delinked:facilitator view"
find OrderNo=fbulbs
___ ENDPROCEDURE (.findbulbs37 _________________________________________________

___ PROCEDURE .NameSearch ______________________________________________________
local firstname, lastname
SelectAll
firstname=""
lastname=""
getscrap "At least one letter from first and last names?"
firstname=clipboard()[1," "][1,-2]
lastname=clipboard()[" ",-1][2,-1]
select Con contains lastname
selectwithin Con contains firstname
message "I've found these many records "+str(info("selected"))
;if info("selected")=info("records")
;openfile "Inquiry List"
;insertbelow
;editcellstop
;endif
window "45 facilitator mailing list:view orders blue box" 
;superobject "seedorders", "FillList"
;superobject "ogsorders", "FillList"
;superobject "treeorders", "FillList"
;superobject "officeorders", "FillList"

___ ENDPROCEDURE .NameSearch ___________________________________________________

___ PROCEDURE .FarmSearch ______________________________________________________
local farm
SelectAll
farm=""
getscrap "What's a part of the group or farm name?"
farm=clipboard()
select Group contains farm
message "I've found these many records "+str(info("selected"))
;if info("selected")=info("records")
;openfile "Inquiry List"
;insertbelow
;editcellstop
;endif
window "45 facilitator mailing list:view orders blue box" 
;superobject "seedorders", "FillList"
;superobject "ogsorders", "FillList"
;superobject "treeorders", "FillList"
;superobject "officeorders", "FillList"

___ ENDPROCEDURE .FarmSearch ___________________________________________________

___ PROCEDURE .update __________________________________________________________
; is this used? jl 8/5/15
; just updated years on files! jl 4/16/18

Numb=«C#»
waswindow=info("windowname")
OpenSecret "45 mailing list"
synchronize
window "Hide This Window"
window waswindow
openfile "&&45 mailing list"
save

window "45seedstally:facilitator view"
synchronize
field OrderNo
sortup
save

window "45ogstally:facilitator view"
synchronize
field OrderNo
sortup
save

;window "45bulbstally:facilitator view"
;synchronize
;field OrderNo
;sortup
;save

;window "45treestally:facilitator view"
;synchronize
;call UPDATE MACRO
;field OrderNo
;sortup
;save

window "45shipping:packages"
synchronize
field «O#»
sortup
save

window "45orders:facilitator view"
synchronize
field OrderNo
sortup
save

window waswindow
;call "listsortcomplete/1"
find «C#»=Numb
;superobject "seedorders", "FillList"
;superobject "ogsorders", "FillList"
;superobject "treeorders", "FillList"
___ ENDPROCEDURE .update _______________________________________________________

___ PROCEDURE (.updatecatalogs _________________________________________________
; is this used? jl 8/5/15


YesNo "Do yo want to update the real mailing list?"
If clipboard()="Yes"
if «C#»>0
Num=«C#»
else place=Con
endif
window "37 mailing list:Secret"
;OpenSecret "37 mailing list"
if Num>0
find Num=«C#»
else 
find Con=place
    if info("found")=0
    YesNo "Open Inquiry List"
        if clipboard()="Yes"
        openfile "Inquiry Lis"
        insertbelow
        field Con
        editcellstop
        endif
    endif
endif    
S=grabdata("37 facilitator mailing list", S)
Bf=grabdata("37 facilitator mailing list", Bf)
T=grabdata("37 facilitator mailing list", T)
;window "Hide This Window"
EndIf
___ ENDPROCEDURE (.updatecatalogs ______________________________________________

___ PROCEDURE listsortcomplete/1 _______________________________________________
;Hide
Field "MAd"
SortUp
Field "City"
SortUp
Field "St"
SortUp
Field "Zip"
SortUp
;Show
;Field «C#»
FirstRecord
Field Con


___ ENDPROCEDURE listsortcomplete/1 ____________________________________________

___ PROCEDURE hidewindow/h _____________________________________________________
Window "Hide This Window"


___ ENDPROCEDURE hidewindow/h __________________________________________________

___ PROCEDURE writeemail/5 _____________________________________________________
; is this used? jl 8/5/15


local mailaddress, mailcopies, mailheader, messageBody
applescript |||
tell application "Finder"
	activate
	tell application "Mail"
		activate
	end tell
end tell
tell application "Panorama"
	activate
end tell
|||
mailaddress=""
mailcopies=""
mailheader=""
messageBody=""
setwindowrectangle
rectanglesize(59, 107, 689, 663), ""
OpenForm "Emailform"
___ ENDPROCEDURE writeemail/5 __________________________________________________

___ PROCEDURE insert below/i ___________________________________________________
InsertBelow
Field «Con»
___ ENDPROCEDURE insert below/i ________________________________________________

___ PROCEDURE getzip/Ω _________________________________________________________
serverlookup "off"
selectall
Numb=0
getscrap "What's the zip?"
Num=val(clipboard())
Find Zip=Num
field MAd
    GetText "What Address", place
        if place≠""   
        find MAd contains place and Zip=Num
        else
            if info("found")=0
            stop
            endif
            beep
         else
         endif
 serverlookup "on"           

___ ENDPROCEDURE getzip/Ω ______________________________________________________

___ PROCEDURE cc rider/ç _______________________________________________________
selectall
waswindow=info("windowname")
serverlookup "off"
GetScrap "enter the customer number"
Find «C#» = val(clipboard())
serverlookup "on"
;if info("windowname")="37 facilitator mailing list:view orders blue box"
;window "37 facilitator mailing list:view orders blue box"
;superobject "ogsorders", "FillList"
;superobject "seedorders", "FillList"
;superobject "treeorders", "FillList"
;superobject "officeorders", "FillList"
;showpage
;endif
;if info("windowname")="37 facilitator mailing list:view 36orders"
;superobject "ogs36orders", "FillList"
;superobject "seed36orders", "FillList"
;superobject "bulbs36orders", "FillList"
;superobject "office36orders", "FillList"
;showpage
;endif

;facilitation still uses this
___ ENDPROCEDURE cc rider/ç ____________________________________________________

___ PROCEDURE noforward/0 ______________________________________________________
S=0
Bf=0
T=0
«M?»=«M?»+"R"
RedFlag="no forward, Trees"
SpareText1=datepattern(today(), "mm/yy")
___ ENDPROCEDURE noforward/0 ___________________________________________________

___ PROCEDURE sameship/2 _______________________________________________________
SAd=MAd
Cit=City
Sta=St
Z=Zip
___ ENDPROCEDURE sameship/2 ____________________________________________________

___ PROCEDURE copy city/3 ______________________________________________________
local address
Hide
UpRecord
arraylinebuild address,¬,"45 facilitator mailing list", City+¬+St+¬+str(Zip)+¬+str(adc)+¶
DownRecord
Show
City=extract(address, ¬,1)
St=extract(address, ¬,2)
Zip=val(extract(address, ¬,3))
adc=val(extract(address, ¬,4))
call "filler/¬"
___ ENDPROCEDURE copy city/3 ___________________________________________________

___ PROCEDURE needseeds/∂ ______________________________________________________
If S=0
S=1
S35=?(S35=0,.01,S35)
endif
«M?»=«M?»+"S"
___ ENDPROCEDURE needseeds/∂ ___________________________________________________

___ PROCEDURE needtrees/® ______________________________________________________
if T=0
T=1
T40=?(T35=0,.01, T40)
endif
«M?»=«M?»+"W"
___ ENDPROCEDURE needtrees/® ___________________________________________________

___ PROCEDURE needbulbs/4 ______________________________________________________
if Bf=0
Bf=1
Bf40=?(Bf40=0,.01,Bf40)
«M?»=«M?»+"Z"
else
«M?»=«M?»+"Z"
endif
___ ENDPROCEDURE needbulbs/4 ___________________________________________________

___ PROCEDURE moved/` __________________________________________________________
;«M?»=«M?»+"U"
;«M?»=«M?»+"E"
«M?»=«M?»+"R"

___ ENDPROCEDURE moved/` _______________________________________________________

___ PROCEDURE filler/¬ _________________________________________________________
Case Zip < 19000  And Zip>1000
S=1
T=1
Bf=1
Case Zip > 40000 And Zip < 60000
S=1
T=1
Bf=0
Case Zip > 97000
S=1
T=1
Bf=0
DefaultCase
S=1
T=0
Bf=0
endcase
if inqcode=""
;inqcode="11fr"
field inqcode 
editcell
endif
if inqcode[3,3]= "t" and T=0
T=1
endif
if inqcode contains "b" and Bf=0
Bf=1
endif
if T=1
«M?»=«M?»+"W"
endif
;if Bf=1
;«M?»=«M?»+"Z"
;endif
if S=1
«M?»=«M?»+"X"
endif
field «C#»
;insertbelow
;field «Con»

___ ENDPROCEDURE filler/¬ ______________________________________________________

___ PROCEDURE inq/œ ____________________________________________________________
field inqcode
editcell
S=1
field Con
___ ENDPROCEDURE inq/œ _________________________________________________________

___ PROCEDURE (Extras) _________________________________________________________

___ ENDPROCEDURE (Extras) ______________________________________________________

___ PROCEDURE refreshlist ______________________________________________________
if «C#»=0
    AlertOk "Sorry, you cannot refresh an inquiry-only record. Put info aside for now."
    stop
    endif

Numb=«C#»
;OpenFile "customer_history:customeractivity"
;find «C#»=Numb
;window "45 facilitator mailing list:view orders blue box"
superobject "seedorders", "filllist"
superobject "ogsorders", "filllist"
superobject "treeorders", "filllist"
superobject "officeorders", "FillList"

OpenFile "45seedstally"
Select «C#»=Numb

OpenFile "45ogstally"
Select «C#»=Numb

;OpenFile "45treestally"
;Select «C#»=Numb

;OpenFile "45bulbstally"
;Select «C#»=Numb

OpenFile "45shipping"
Select «C#»=str(Numb)

window "45 facilitator mailing list:view orders blue box"
___ ENDPROCEDURE refreshlist ___________________________________________________

___ PROCEDURE refill boxes/r ___________________________________________________
if «C#»=0
    AlertOk "Sorry, you cannot refresh an inquiry-only record. Put info aside for now."
    stop
    endif

;superobject "oldseedorders", "filllist"
superobject "ogsorders", "filllist"
superobject "treeorders", "filllist"
superobject "seedorders", "filllist"
superobject "bulborders", "filllist"
superobject "officeorders", "filllist"


___ ENDPROCEDURE refill boxes/r ________________________________________________

___ PROCEDURE close file/∑ _____________________________________________________
closefile 
___ ENDPROCEDURE close file/∑ __________________________________________________

___ PROCEDURE window ___________________________________________________________
Num=  info("WindowBox") 
message  Num
___ ENDPROCEDURE window ________________________________________________________

___ PROCEDURE check ____________________________________________________________
GetScrap "vat iz de zip code?"
Find Zip = val(clipboard())
message info("found")
___ ENDPROCEDURE check _________________________________________________________

___ PROCEDURE Find Return ______________________________________________________
Select Con Contains Chr(13)
SelectAdditional Group Contains Chr(13)
SelectAdditional MAd Contains Chr(13)
SelectAdditional City Contains Chr(13)
SelectAdditional St Contains Chr(13)
___ ENDPROCEDURE Find Return ___________________________________________________

___ PROCEDURE (lookupcustomers _________________________________________________
serverlookup "off"
select «C#»=lookup("30seedspatronage","C#",«C#»,"C#",0,0)
serverlookup "on"

___ ENDPROCEDURE (lookupcustomers ______________________________________________

___ PROCEDURE reselect _________________________________________________________
selectall
SelectDuplicates MAd+City
SelectWithin inqcode  NOTCONTAINS "comp"
___ ENDPROCEDURE reselect ______________________________________________________

___ PROCEDURE deleter __________________________________________________________
NoUndo
Hide
GetScrapOK "what code?"
Select inqcode contains clipboard()
SelectWithin «C#» = 0
Show
___ ENDPROCEDURE deleter _______________________________________________________

___ PROCEDURE forceunlock ______________________________________________________
forceunlockrecord
___ ENDPROCEDURE forceunlock ___________________________________________________

___ PROCEDURE .Seed_comments ___________________________________________________
Openfile "45SeedsComments linked"
GoForm "facilitator view"


___ ENDPROCEDURE .Seed_comments ________________________________________________

___ PROCEDURE .OGS comments ____________________________________________________
Openfile "45ogscomments.linked"
GoForm "facilitator view"

___ ENDPROCEDURE .OGS comments _________________________________________________

___ PROCEDURE .seedspecs _______________________________________________________
Openfile "SEEDSPECS"
GoForm "facilitator readonly"
call "finditem/1"
___ ENDPROCEDURE .seedspecs ____________________________________________________

___ PROCEDURE (.find orders ____________________________________________________
; is this used? jl 8/5/15


;openfile "34treestally"
;selectall
;select «C#»=ID
;if  info("Empty") 
;    if ono=0
;Message "No orders found for this customer in trees orders."
;    endif
;endif


openfile "37ogstally"
selectall
select «C#»=ID
if  info("Empty") 
    if ono=0
Message "No orders found for this customer in OGS tally."
    endif
endif

openfile  "37seedstally"
selectall
select «C#»=ID
if  info("Empty") 
    if ono=0
Message "No orders found for this customer in seeds tally."
endif
endif

;openfile "37bulbstally"
;selectall
;select «C#»=ID
;if  info("Empty") 
;    if ono=0
;Message "No orders found for this customer in bulbs tally."
;endif
;endif

___ ENDPROCEDURE (.find orders _________________________________________________

___ PROCEDURE .O#search ________________________________________________________
GetScrap "What is the order number?"
ono=val(clipboard())
window "45orders:facilitator view"
SelectAll
find OrderNo=ono
Numb=«C#»

window "45 facilitator mailing list:view orders blue box"
SelectAll
find «C#»=Numb
superobject "seedorders", "filllist"
superobject "ogsorders", "filllist"
superobject "treeorders", "filllist"
superobject "officeorders", "filllist"

OpenFile "45ogstally"
SelectAll
Select «C#»=Numb

OpenFile "45seedstally"
SelectAll
Select «C#»=Numb

;OpenFile "45treestally"
;SelectAll
;Select «C#»=Numb

;OpenFile "45bulbstally"
;SelectAll
;Select «C#»=Numb

OpenFile "45shipping"
SelectAll
Select «C#»=str(Numb)

window "45 facilitator mailing list:view orders blue box"
___ ENDPROCEDURE .O#search _____________________________________________________

___ PROCEDURE (editInquiry _____________________________________________________
; is this used? jl 8/5/15



global icon, igroup, imad, icity, ist, izip
if «C#»≥1
AlertOk "This procedure is for non-customers only."
stop
endif
stop
icon=Con
igroup=Group
imad=MAd
icity=City
ist=St
izip=Zip




Hide
Open "34 mailing list"
window "34 mailing list:Secret"
Select


Con=grabdata("34 facilitator mailing list", Con)
Group=grabdata("34 facilitator mailing list", Group)
MAd=grabdata("34 facilitator mailing list", MAd)
City=grabdata("34 facilitator mailing list", City)
St=grabdata("34 facilitator mailing list", St)
Zip=grabdata("34 facilitator mailing list", Zip)
adc=grabdata("34 facilitator mailing list", adc)
phone=grabdata("34 facilitator mailing list", phone)
email=grabdata("34 facilitator mailing list", email)
S=grabdata("34 facilitator mailing list", S)
Bf=grabdata("34 facilitator mailing list", Bf)
T=grabdata("34 facilitator mailing list", T)
;window "Hide This Window"
window "34 facilitator mailing list:Edit Record"
CloseWindow
___ ENDPROCEDURE (editInquiry __________________________________________________

___ PROCEDURE (refresh35 _______________________________________________________
; is this used? jl 8/5/15


if «C#»>0
Numb=«C#» 
window "customer_history:customeractivity"
find «C#»=Numb
window "37 facilitator mailing list:view 36orders"
superobject "seed36orders", "filllist"
superobject "ogs36orders", "filllist"
superobject "bulb36orders", "filllist"
superobject "office36orders", "FillList"

OpenFile "36seedstally"
Select «C#»=Numb

OpenFile "36ogstally"
Select «C#»=Numb

OpenFile "36bulbstally"
Select «C#»=Numb

;OpenFile "36treestally"
;Select «C#»=Numb

;OpenFile "34moosetally"
;Select «C#»=Numb

OpenFile "36shipping"
Select «C#»=str(Numb)

window "37 facilitator mailing list:view 36orders"
endif
___ ENDPROCEDURE (refresh35 ____________________________________________________

___ PROCEDURE (.findogs35 ______________________________________________________
global fogs
fogs=0
fogs=val(extract(35ogsorders,"  ",1))
window "35ogstally:facilitator view"
select OrderNo=fogs
___ ENDPROCEDURE (.findogs35 ___________________________________________________

___ PROCEDURE (.findseeds35 ____________________________________________________
global fseeds
fseeds=0
fseeds=val(extract(35seedorders,"  ",1))
window "35seedstally:facilitator view"
find OrderNo=fseeds
___ ENDPROCEDURE (.findseeds35 _________________________________________________

___ PROCEDURE exportprocedures _________________________________________________
global facilitatorprocedures, newprocedures
SaveALLProcedures "", facilitatorprocedures

___ ENDPROCEDURE exportprocedures ______________________________________________

___ PROCEDURE closefile ________________________________________________________
closefile

___ ENDPROCEDURE closefile _____________________________________________________

___ PROCEDURE open history _____________________________________________________
Numb=«C#»
waswindow = info("windowname")
OpenFile "customer_history"
GoForm "customeractivity"
    zoomwindow 783,0,400,647,""
Select «C#»=Numb
window waswindow

___ ENDPROCEDURE open history __________________________________________________

___ PROCEDURE open inquiry _____________________________________________________
openfile "Inquiry List"
lastrecord
insertbelow
___ ENDPROCEDURE open inquiry __________________________________________________

___ PROCEDURE test _____________________________________________________________
Openfile "45ogstally"
    zoomwindow 51,1036,899,870,""
    field OrderNo
    sortup
    save
Openfile "45seedstally"
   zoomwindow 83,959,951,932,""
    field OrderNo
    sortup
    save
___ ENDPROCEDURE test __________________________________________________________

___ PROCEDURE huh? _____________________________________________________________
Openfile "45orders"
    goform "facilitator view"
    zoomwindow 25,1159,764,760,""
    save
___ ENDPROCEDURE huh? __________________________________________________________

___ PROCEDURE .ZeroOutCatalogs _________________________________________________
if RedFlag contains "no catalog"
field S
S=0
field Bf
Bf=0
field T
T=0
endif

if RedFlag contains "no seeds catalog" 
field S
S=0
endif

if RedFlag contains "no bulbs catalog"
field Bf
Bf=0
endif

if RedFlag contains "no trees catalog"
field T
T=0
endif


___ ENDPROCEDURE .ZeroOutCatalogs ______________________________________________

___ PROCEDURE .SameAsMailing ___________________________________________________
field SAd
SAd=MAd
field Cit
Cit=City
field Sta
Sta=St
field Z
Z=Zip
___ ENDPROCEDURE .SameAsMailing ________________________________________________

___ PROCEDURE FindBadAddCanadians ______________________________________________
local vCanada
vCanada="NL
PE
NS
NB
QC
ON
MB
SK
AB
BC
YT
NT
NU"

select vCanada contains St
___ ENDPROCEDURE FindBadAddCanadians ___________________________________________

___ PROCEDURE GetSource ________________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros from " +info("databasename")+" are saved to your clipboard!"
___ ENDPROCEDURE GetSource _____________________________________________________
