___ PROCEDURE .Initialize ______________________________________________________
Global New, Num, Numb, enter, place, ID, CC, ED, credit, orders, ship, Flag, waswindow
fileglobal searchcust, findcust, getcust, searchname, findname, getname
expressionstacksize 75000000
Num=0
findname=""
searchname=""
searchcust=""
;window "45 mailing list:addresschecker"
;field Cit
gosheet
noshow
call "listsortcomplete/1"
endnoshow
waswindow = info("DatabaseName") 
openfile "fcmadc"
makesecret
;window waswindow
;YesNo "Do you want to open all the files?"
;if clipboard()="No"
;stop
;endif
Openfile "customer_hIstory"
Openfile "45orders"
;Openfile "44orders"
//Openfile "Customer#"

window  waswindow
openfile "ZipCodeList"
makesecret
window waswindow
___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE .addmember _______________________________________________________
if info("files") notcontains "members"
openfile "members"
else
window "members"
endif
lastrecord
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Con=grabdata("45 mailing list", Con)
Group=grabdata("45 mailing list", Group)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
SAd=grabdata("45 mailing list", SAd)
Cit=grabdata("45 mailing list", Cit)
Sta=grabdata("45 mailing list", Sta)
Z=grabdata("45 mailing list", Z)
phone=grabdata("45 mailing list", phone)
email=grabdata("45 mailing list", email)
inqcode=grabdata("45 mailing list", inqcode)
«Mem?»=grabdata("45 mailing list", «Mem?»)
windowtoback "members"
window "45 mailing list"
___ ENDPROCEDURE .addmember ____________________________________________________

___ PROCEDURE NewRecord ________________________________________________________
if info("trigger") ="New.Return"
insertrecord
else
Synchronize
beep
endif
___ ENDPROCEDURE NewRecord _____________________________________________________

___ PROCEDURE .DeleteRecord ____________________________________________________
field «C#»
copycell
YesNo "delete this record?"
if clipboard()="No"
stop
endif
deleterecord
if val(clipboard())>0
window "customer_history:customeractivity"
if «C#»≠val(clipboard())
find «C#»=val(clipboard())
if info("found")=0
window "45 mailing list"
stop
endif
endif
yesno "delete this record?"
if clipboard()="Yes"
deleterecord
endif
window "45 mailing list"
endif
if info("selected")<info("records")
downrecord
endif
___ ENDPROCEDURE .DeleteRecord _________________________________________________

___ PROCEDURE .KeyDown _________________________________________________________
local KeyStroke
KeyStroke=info("trigger")[5,-1]
case KeyStroke=chr(31)
;if info("formName")="customer history" and info("selected")<info("records")
;downrecord
;if  info("Files") notcontains "inquiries&changes"
;openfile "inquiries@changes"
;endif
;window "inquiries&changes:change"
;downrecord
;window "45 mailing list:customer history"
;else
;downrecord
;endif
if info("formname")="addresschecker"
downrecord
window "45orders:seedsinput"
downrecord
window "45 mailing list:addresschecker"
endif
defaultcase
    key info("modifiers"), KeyStroke
    endcase
___ ENDPROCEDURE .KeyDown ______________________________________________________

___ PROCEDURE .customer ________________________________________________________
OpenFile "45orders"
Case Numb>300000 And Numb<400000
    GoForm "ogsinput"
Case Numb>400000 And Numb<500000
    GoForm "treesinput"
Case Numb>50000 And Numb<60000
    GoForm "bulbsinput"
Case Numb>600000 And Numb<700000
    GoForm "seedsinput"
Case Numb>700000
    GoForm "seedsinput"
EndCase
call ".customerfill"

___ ENDPROCEDURE .customer _____________________________________________________

___ PROCEDURE .findcustomer ____________________________________________________
if info("trigger")="Button.Find Customer"
if info("files") contains "45orders"
goto continue
else
opensecret "45orders"
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

___ ENDPROCEDURE .findcustomer _________________________________________________

___ PROCEDURE .tab _____________________________________________________________
If info("trigger")="Key.Return"
Field «C#»
EndIf
___ ENDPROCEDURE .tab __________________________________________________________

___ PROCEDURE .holdit __________________________________________________________
global VC
VC=«C#»
If «C#»=0
field Group
editcellstop
endif
If info("changes")≠0
CancelOk "Do you mean to change the C#?"
If clipboard()="Cancel"
«C#»=VC
EndIf
EndIf


___ ENDPROCEDURE .holdit _______________________________________________________

___ PROCEDURE .closewindow _____________________________________________________
CloseWindow
___ ENDPROCEDURE .closewindow __________________________________________________

___ PROCEDURE .MAd _____________________________________________________________
if «C#»=0
stop
endif
;if info("trigger")="Key.Tab"
;right
;endif
if info("trigger")="Key.Return"
Num=«C#»
openfile "customer_history"
find «C#»=Num
if info("found")=0
insertrecord
«C#»=Num
Con=grabdata("45 mailing list", Con)
Group=grabdata("45 mailing list", Group)
endif
;Con=grabdata("45 mailing list", Con)
;Group=grabdata("45 mailing list", Group)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
SpareText2=grabdata("45 mailing list", SpareText2)
Notes=replace(Notes, "Bad Address","")
Num=0
endif
window "45 mailing list"
___ ENDPROCEDURE .MAd __________________________________________________________

___ PROCEDURE .newzip __________________________________________________________
fileglobal listzip, thiszip, findher, findzip, findcity, newcity, findname,findname1, findname2, thisname, firstname, lastname
serverlookup "off" 
;waswindow=info("windowname")
listzip=""
thiszip=""
newcity=""
again:
findher=aray


supergettext findher, {caption="Enter Address.Zip" height=100 width=400 captionfont=Times captionsize=14 captioncolor="cornflowerblue"
    buttons="Find;Redo;Cancel"}
    if info("dialogtrigger") contains "Find"
        findzip=extract(findher,".",2)
        findzip=strip(findzip)
            if length(findzip)=4
            findzip="0"+findzip
            endif
        findcity=extract(findher,".",1)
        liveclairvoyance findzip,listzip,¶,"","45 mailing list",pattern(Zip,"#####"),"=",str(«C#»)+¬+rep(" ",7-length(str(«C#»)))+Con+rep(" ",max(20-length(Con),1))+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####"),0,0,""
        arraysubset listzip, listzip, ¶, import() contains findcity
            if listzip=""
            goto lastzip
            endif
    
        if arraysize(listzip,¶)=1
        find MAd contains findcity and pattern(Zip,"#####") contains findzip
        AlertYesNo "Enter this one?"
            if info("dialogtrigger") contains "Yes"
           goto lastline
            ;stop
            else
            AlertOkCancel "Try by zipcode?"
                if info("dialogtrigger") contains "OK"
                call "getzip/Ω"
                endif
             endif
           endif
    endif
    
    if info("dialogtrigger") contains "Redo"
    findher=""
    goto again
    endif
    
    if info("dialogtrigger") contains "Cancel"
    window waswindow
    stop
    endif


superchoicedialog listzip, thiszip, {height=400 width=800 font=Courier caption="Click on one and then hit OK or New for new entry" 
        captionfont=Times captionsize=12 captioncolor=red size=14 buttons="OK:100;Try Name:150;Cancel:100"}
if info("dialogtrigger") contains "OK"
    find «C#» = val(strip(extract(thiszip, ¬,1))) and MAd=extract(thiszip, ¬,3) and City contains extract(thiszip, ¬,4)
    ;;find MAd=extract(thiszip, ¬,2) and City contains extract(thiszip, ¬,3)
    showpage

    call "enter/e"
endif

if info("dialogtrigger") contains "Try Name"
    goto tryname
    gettext "Which town?", newcity
    if newcity≠""
        find Z=val(findzip) and City contains newcity
        insertbelow
    else
        find Zip=val(findzip)
        insertbelow
    endif
endif
showpage
serverlookup "on"

tryname:
    firstname=""
    lastname=""
    findname=""
    findname1=""
    findname2=""
    findname=rayj
    supergettext findname, {caption="Enter First and Last Name" height=100 width=400 captionfont=Times captionsize=14 captioncolor="limegreen"
    buttons="Find;Redo;Cancel"}
    firstname=extract(findname," ",1)
     lastname=extract(findname," ",2)
    if info("dialogtrigger") contains "Find"
        liveclairvoyance lastname,findname1,¶,"","45 mailing list",Con,"contains",Con+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")+¬+phone,0,0,""
        message findname1
    endif
    
    if info("dialogtrigger") contains "Redo"
        goto tryname
    endif
    
    if info("dialogtrigger") contains "Cancel"
        stop
    endif
    
    arraysubset findname1,findname1,¶,import() contains firstname
    if arraysize(findname1,¶)=1
        find Con contains firstname and Con contains lastname
        AlertYesNo "Enter this one?"
        if info("dialogtrigger") contains "Yes"
            call "enter/e"
            stop
        else
            lastzip:
            getscrap "What zip code?"
            find Zip=val(clipboard())
            insertbelow
            stop
        endif
    endif
    superchoicedialog findname1,thisname, {height=400 width=500 font=Helvetica caption="Click on one and then hit OK or New for new entry" 
        captionfont=Times captionsize=12 captioncolor=green size=14 buttons="OK:100;New:100;Cancel:100"}
     if info("dialogtrigger") contains "OK"
        find Con contains extract(thisname, ¬,1) and City contains extract(thisname, ¬,3)
        call "enter/e"
     endif
     
     if info("dialogtrigger") contains "New"
        gettext "Which town?", newcity
        if newcity≠""
            window "45 mailing list"
            find Z=val(findzip) and City contains newcity
            insertbelow
        else
            find Zip=val(findzip)
            insertbelow
        endif
    endif
    
    showpage
    serverlookup "on"
    stop
    lastline:
    serverlookup "on"
    call "enter/e"


___ ENDPROCEDURE .newzip _______________________________________________________

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

___ PROCEDURE .customerhistory _________________________________________________
         if «C#»>0
            Numb=«C#»
            window "customer_history:customer history"
            find «C#»=Numb
            window waswindow
            endif
___ ENDPROCEDURE .customerhistory ______________________________________________

___ PROCEDURE .tab1 ____________________________________________________________
;if info("fieldname")="C#"
style "record black"
;endif
Numb=0
if «C#»>0
Numb=«C#»
window "customer_history:secret"
Find «C#» = Numb
if info("found")=0
OpenSheet
lastrecord
insertbelow
endif
«C#»=grabdata("45 mailing list", «C#»)
Con=grabdata("45 mailing list", Con)
Group=grabdata("45 mailing list", Group)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
«SpareText2»=grabdata("45 mailing list", «SpareText2»)
;closewindow
endif
window "customer_history:customeractivity"

window "45 mailing list"

___ ENDPROCEDURE .tab1 _________________________________________________________

___ PROCEDURE .member __________________________________________________________
if «Mem?»≠"Y"
stop
endif
if «C#»=0
Code= "45mem"
openfile "Customer#"
call "newnumber"
window "45 mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
if S=0
call "filler/¬"
endif
field inqcode
window "customer_history"
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow
window "45 mailing list"
field Notes
field «inqcode»
endif
Num=«C#»
window "customer_history:secret"
find «C#»=Num
if NewMember notcontains "joined"
NewMember=NewMember+"joined on "+datepattern(today(),"mm/dd/yy")
endif
field «Equity»
getscrap "How much equity?"
«Equity»=val(clipboard())
window "45 mailing list"
call .addmember
___ ENDPROCEDURE .member _______________________________________________________

___ PROCEDURE listsortcomplete/1 _______________________________________________
Hide
Field "MAd"
SortUp
Field "City"
SortUp
Field "St"
SortUp
Field "Zip"
SortUp
Show
Field «C#»



___ ENDPROCEDURE listsortcomplete/1 ____________________________________________

___ PROCEDURE hidewindow/h _____________________________________________________
Window "Hide This Window"


___ ENDPROCEDURE hidewindow/h __________________________________________________

___ PROCEDURE (Entering) _______________________________________________________

___ ENDPROCEDURE (Entering) ____________________________________________________

___ PROCEDURE createcustomer/µ _________________________________________________
if «C#»≠0
stop 
endif
Code= "I45s"
openfile "Customer#"
call "newnumber"
window "45 mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
If inqcode=""
Field inqcode
inqcode=?(inqcode contains "17", inqcode[3,-1], inqcode)
EditCell
field «C#»
EndIf
if S=0
call "filler/¬"
endif
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow
window "45 mailing list"
___ ENDPROCEDURE createcustomer/µ ______________________________________________

___ PROCEDURE cc rider/ç _______________________________________________________
waswindow=info("windowname")
serverlookup "off"
if info("trigger")="Button.Find Customer #"
if info("files") contains "45orders"
goto continue
else
opensecret "45orders"
endif
continue:
endif
NoUndo
GetScrap "enter the customer number"
Find «C#» = val(clipboard())
window "customer_history:secret"
Find «C#» = val(clipboard())
window waswindow
if info("fieldname") notcontains "Mem?"
field «C#»
field MAd
endif
serverlookup "on"
___ ENDPROCEDURE cc rider/ç ____________________________________________________

___ PROCEDURE entering/√ _______________________________________________________
getscrap "Next! (all 6 digits, please)"
Numb=val(clipboard())
OpenFile "45orders"
ReSynchronize
field OrderNo
sortup
Case Numb<300000
    GoForm "seedsinput"
Case Numb>300000 And Numb<400000
    GoForm "ogsinput"
Case Numb>400000 And Numb<500000
    GoForm "treesinput"
Case Numb>50000 And Numb<60000
    GoForm "bulbsinput"
Case Numb>600000 And Numb<700000
    GoForm "mtinput"
Case Numb>700000
    GoForm "seedsinput"
EndCase
waswindow=info("windowname")
window waswindow
find OrderNo=val(clipboard())
field «C#»


___ ENDPROCEDURE entering/√ ____________________________________________________

___ PROCEDURE enter/e __________________________________________________________
Num=«C#»
If Num=0
;Message "You must create a Customer#"
                if ono>300000 and ono<400000
                    call "ogsity/ø"
                endif
                if ono>50000 and ono<60000
                    call "bulbous/∫"
                endif
                if ono>400000 and ono<500000
                    call "treed/†"
                endif
                 if ono>700000
                    call "seedy/ß"
                endif
                if ono>600000 and ono<700000
                    call "moosed/µ"
                endif
endif
case waswindow contains "bulbs"
Bf=?(Bf=0,1,Bf)
case waswindow contains "tree"
T=?(T=0,1,T)
case waswindow contains "seed"
S=?(S=0,1,S)
case waswindow contains "ogs"
S=?(S=0,1,S)
case waswindow contains "mt"
S=?(S=0,1,S)
endcase
window waswindow
«C#»=Num
«C#Text»=str(Num)
case info("formname") = "ogsinput" and OrderNo>320000 and OrderNo<400000
    if Con≠""
        stop
    endif
case info("formname") = "seedsinput" and OrderNo>710000
    if Con≠""
        stop
    endif
case info("formname") = "treesinput" and OrderNo>420000 and OrderNo<500000
    if Con≠""
        stop
    endif
case info("formname") = "mtinput" and OrderNo>620000 and OrderNo<700000
    if Con≠""
        stop
    endif
case info("formname") = "bulbsinput" and  OrderNo>54000 and OrderNo<60000
    if Con≠""
        stop
    endif
endcase

call ".customerfill"
___ ENDPROCEDURE enter/e _______________________________________________________

___ PROCEDURE sameship/2 _______________________________________________________
if «C#»=0
stop
endif
if MAd contains "PO " and Z=Zip
stop
else if MAd contains "PO " and Z≠Zip
SAd=""
else SAd=MAd
endif
endif
Cit=City
Sta=St
Z=Zip
___ ENDPROCEDURE sameship/2 ____________________________________________________

___ PROCEDURE seedy/ß __________________________________________________________
if «C#»≠0
stop 
endif
Code= "I45s"
openfile "Customer#"
call "newnumber"
window "45 mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
If inqcode=""
Field inqcode
inqcode=?(inqcode contains "17", inqcode[3,-1], inqcode)
EditCell
field «C#»
EndIf
if S=0
call "filler/¬"
endif
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow
window "45 mailing list"
Call "enter/e"

___ ENDPROCEDURE seedy/ß _______________________________________________________

___ PROCEDURE ogsity/ø _________________________________________________________
if «C#»≠0
stop 
endif
Code="I45o"
openfile "Customer#"
call "newnumber"
window "45 mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
If inqcode=""
Field inqcode
EditCell
field «C#»
EndIf
if S=0
call "filler/¬"
endif
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow
window "45 mailing list"
Call "enter/e"



___ ENDPROCEDURE ogsity/ø ______________________________________________________

___ PROCEDURE moosed/µ _________________________________________________________
if «C#»≠0
stop 
endif
Code= "I45m"
openfile "Customer#"
call "newnumber"
window "45 mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
If inqcode=""
Field inqcode
EditCell
field «C#»
EndIf
if S=0
call "filler/¬"
endif
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow
window "45 mailing list"
Call "enter/e"




___ ENDPROCEDURE moosed/µ ______________________________________________________

___ PROCEDURE treed/† __________________________________________________________
if «C#»≠0
stop 
endif
Code="I45t"
openfile "Customer#"
call "newnumber"
window "45 mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
If inqcode=""
Field inqcode
EditCell
field «C#»
EndIf
if S=0
call "filler/¬"
endif
if T=0
T=1
endif
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow
window "45 mailing list"
Call "enter/e"

___ ENDPROCEDURE treed/† _______________________________________________________

___ PROCEDURE bulbous/∫ ________________________________________________________
if «C#»≠0
stop 
endif
Code= "I45b"
openfile "Customer#"
call "newnumber"
window "45 mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
If inqcode=""
Field inqcode
EditCell
field «C#»
EndIf
if S=0
call "filler/¬"
endif
If Bf=0
Bf=1
endif
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata("45 mailing list", «C#»)
Group=grabdata("45 mailing list", Group)
Con=grabdata("45 mailing list", Con)
MAd=grabdata("45 mailing list", MAd)
City=grabdata("45 mailing list", City)
St=grabdata("45 mailing list", St)
Zip=grabdata("45 mailing list", Zip)
Email=grabdata("45 mailing list", email)
SpareText2=grabdata("45 mailing list", SpareText2)
;CloseWindow
window "45 mailing list"
Call "enter/e"

___ ENDPROCEDURE bulbous/∫ _____________________________________________________

___ PROCEDURE writeemail/5 _____________________________________________________
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

___ PROCEDURE (Inquiries) ______________________________________________________


___ ENDPROCEDURE (Inquiries) ___________________________________________________

___ PROCEDURE getzip/Ω _________________________________________________________
fileglobal listzip, thiszip, findher, findzip, findaddress, newcity, arraynumb,sortzip, sortcity
serverlookup "off" 
sortzip=""
sortcity=""
arraynumb=0
listzip=""
thiszip=""
newcity=""
again:
findher=""
waswindow=info("windowname")



supergettext findher, {caption="Enter Address.Zip or .Zip to find everyone" height=100 width=400 captionfont=Times captionsize=14 captioncolor="cornflowerblue"
    buttons="Find;Redo;Cancel"}
    if info("dialogtrigger") contains "Find"
        findzip=extract(findher,".",2)
        findzip=strip(findzip)
        if length(findzip)=4
            findzip="0"+findzip
        endif
        findaddress=extract(findher,".",1)
        liveclairvoyance findzip,listzip,¶,"","45 mailing list",pattern(Zip,"#####"),"=",str(«C#»)+¬+rep(" ",7-length(str(«C#»)))+Con+rep(" ",max(20-length(Con),1))+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####"),0,0,""
        if findaddress=""
            sortzip=listzip
        else
            arraysubset listzip, listzip, ¶, import() contains findaddress
            loop arraynumb=arraynumb+1
            stoploopif arraynumb>arraysize(listzip,¶)
                sortzip=sortzip+?(extract(extract(listzip,¶,arraynumb),¬,3) contains findaddress,extract(listzip,¶,arraynumb)+¶,"")
            until arraynumb=arraysize(listzip,¶)+1
        endif
        if info("found")=0
            beep
            NoYes "No one in that zip. Try another?"
            findher=""
            If clipboard()="Yes"
                goto again
            else
                insertbelow
                stop
            endif
        endif
        if arraysize(listzip,¶)=1
            find MAd contains findaddress and pattern(Zip,"#####") contains findzip
            field MAd
            if «C#»>0
                CC=«C#»
                window "customer_history:customeractivity"
                find «C#»=CC
                window waswindow
            endif
            stop
        endif
    endif
    if info("dialogtrigger") contains "Redo"
        findher=""
        goto again
    endif
    if info("dialogtrigger") contains "Cancel"
        stop
    endif

superchoicedialog sortzip, thiszip, {height=400 width=800 font=Courier caption="Click on one and then hit Choose or New for new entry" 
    captionfont=Times captionsize=12 captioncolor=red size=16 buttons="OK:100;New:100;Cancel:100"}
if info("dialogtrigger") contains "OK"
    
    find «C#» = val(strip(extract(thiszip, ¬,1))) and MAd=extract(thiszip, ¬,3) and City contains extract(thiszip, ¬,4)
    ;;find MAd=extract(thiszip, ¬,3) and City contains extract(thiszip, ¬,4)
    field MAd
    if «C#»>0
        CC=«C#»
        window "customer_history:customeractivity"
        find «C#»=CC
        window waswindow
    endif
endif
if info("dialogtrigger") contains "New"
    arrayfilter sortzip, sortcity, ¶, extract(extract(sortzip,¶,seq()),¬,4)
    arraydeduplicate sortcity,  sortcity, ¶

    if (findaddress≠"" and arraysize(sortcity,¶)>2) or (findaddress="" and arraysize(sortcity,¶)>1)
        gettext "Which town?", newcity
    endif
    if newcity≠""
        find Z=val(findzip) and City contains newcity
        insertbelow
        field Con
    else
        find Zip=val(findzip)
        insertbelow
        field Con
    endif
endif
showpage
serverlookup "on"
___ ENDPROCEDURE getzip/Ω ______________________________________________________

___ PROCEDURE insert below/i ___________________________________________________
InsertBelow
Field «Con»
___ ENDPROCEDURE insert below/i ________________________________________________

___ PROCEDURE noforward/0 ______________________________________________________
Numb=«C#»
S=0
Bf=0
T=0
«M?»=«M?»+"E"
RedFlag="no forward, bulbs"+¶+RedFlag
SpareText1=datepattern(today(), "mm/yy")
if «C#»=0
field «C#»
stop
endif
waswindow=info("windowname")
window "customer_history:customeractivity"
Find «C#» = Numb
Notes=?(Notes="", "Bad address", Notes+¶+"Bad Address")
window waswindow
___ ENDPROCEDURE noforward/0 ___________________________________________________

___ PROCEDURE no mail/µ ________________________________________________________
Numb=«C#»
S=0
Bf=0
T=0
«M?»=«M?»+"R"
RedFlag="no mail receptacle"+¶+RedFlag
SpareText1=datepattern(today(), "mm/yy")
if «C#»=0
field «C#»
stop
endif
waswindow=info("windowname")
window "customer_history:customeractivity"
Find «C#» = Numb
Notes=?(Notes="", "No mail receptacle", Notes+¶+"Bad Address")
window waswindow
___ ENDPROCEDURE no mail/µ _____________________________________________________

___ PROCEDURE tempaway/y _______________________________________________________
Numb=«C#»
S=0
Bf=0
T=0
«M?»=«M?»+"E"
RedFlag="temporarily away"+¶+RedFlag
SpareText1=datepattern(today(), "mm/yy")
if «C#»=0
field «C#»
stop
endif
waswindow=info("windowname")
window "customer_history:customeractivity"
Find «C#» = Numb
Notes=?(Notes="" or Notes="Bad address", "Away", Notes+¶+"Away")
window waswindow
___ ENDPROCEDURE tempaway/y ____________________________________________________

___ PROCEDURE copy city/3 ______________________________________________________
local address
UpRecord
arraylinebuild address,¬,"45 mailing list", City+¬+St+¬+str(Zip)+¬+str(adc)+¶
DownRecord
City=extract(address, ¬,1)
St=extract(address, ¬,2)
Zip=val(extract(address, ¬,3))
adc=val(extract(address, ¬,4))
call "filler/¬"
___ ENDPROCEDURE copy city/3 ___________________________________________________

___ PROCEDURE filler/¬ _________________________________________________________
Case Zip < 19000  And Zip>1000
S=1
T=1
Bf=1
Case (Zip > 43000 And Zip < 46000) or (Zip > 48000 And Zip < 50000) or (Zip > 53000 And Zip < 57000) or Zip>97000
S=1
T=1
Bf=0
DefaultCase
S=1
T=0
Bf=0
endcase
If inqcode=""
field inqcode
editcell
endif
if inqcode[3,3]= "t" and T=0
T=1
endif
if inqcode contains "b" and Bf=0
Bf=0
endif
;if T=1
;«M?»="W"
;endif
;if Bf=1
;«M?»=«M?»+"Z"
;endif
;if S=1
;«M?»=«M?»+"X"
;endif
field «C#»
___ ENDPROCEDURE filler/¬ ______________________________________________________

___ PROCEDURE moved/` __________________________________________________________
;«M?»=«M?»+"U"
«M?»=«M?»+"E"
;«M?»=«M?»+"R"
;if inqcode contains "onl"
;S=0
;Bf=0
;T=0
;endif
SpareText1=datepattern(today(),"mm/yy")
call "filler/¬"
___ ENDPROCEDURE moved/` _______________________________________________________

___ PROCEDURE inq/œ ____________________________________________________________
field inqcode
editcell
S=1
field Con
___ ENDPROCEDURE inq/œ _________________________________________________________

___ PROCEDURE needseeds/6 ______________________________________________________
;If S=0
;S=1
;endif
«M?»=«M?»+"S"
___ ENDPROCEDURE needseeds/6 ___________________________________________________

___ PROCEDURE needtrees/® ______________________________________________________
if T=0
T=1
endif
«M?»=«M?»+"W"
___ ENDPROCEDURE needtrees/® ___________________________________________________

___ PROCEDURE needbulbs/4 ______________________________________________________
if Bf=0
Bf=1
«M?»=«M?»+"Z"
endif
___ ENDPROCEDURE needbulbs/4 ___________________________________________________

___ PROCEDURE (Extras) _________________________________________________________


___ ENDPROCEDURE (Extras) ______________________________________________________

___ PROCEDURE window ___________________________________________________________
Num=  info("WindowBox") 
message  Num
___ ENDPROCEDURE window ________________________________________________________

___ PROCEDURE deletem __________________________________________________________
lastrecord
loop
repeatloopif «Mem?»="Y"
deleterecord
until info("selected")=1
field Con
copy
selectall
find Con=clipboard()
deleterecord
___ ENDPROCEDURE deletem _______________________________________________________

___ PROCEDURE checktowns _______________________________________________________
select Cit contains chr(13) or Cit endswith chr(32) or Cit contains chr(44)

___ ENDPROCEDURE checktowns ____________________________________________________

___ PROCEDURE deleter __________________________________________________________
NoUndo
Hide
GetScrapOK "what code?"
Select inqcode contains clipboard()
SelectWithin «C#» = 0
Show
___ ENDPROCEDURE deleter _______________________________________________________

___ PROCEDURE check ____________________________________________________________
GetScrap "vat iz de zip code?"
Find Zip = val(clipboard())
message info("found")
___ ENDPROCEDURE check _________________________________________________________

___ PROCEDURE inq&change _______________________________________________________
WindowBox "22 324 219 915"
openfile "37inquiries&changes"
window "change"
select «C#»>0
field «C#»
sortup
Window "37mailing list:customer history"
Select lookupselected("37inquiries&changes","C#",«C#»,"C#",0,0)
field «C#»
sortup 
___ ENDPROCEDURE inq&change ____________________________________________________

___ PROCEDURE Find Return ______________________________________________________
Select Con Contains Chr(13)
SelectAdditional Group Contains Chr(13)
SelectAdditional MAd Contains Chr(13)
SelectAdditional City Contains Chr(13)
SelectAdditional St Contains Chr(13)
___ ENDPROCEDURE Find Return ___________________________________________________

___ PROCEDURE lookupcustomers __________________________________________________
serverlookup "off"
select «C#»=lookup("30seedspatronage","C#",«C#»,"C#",0,0)
serverlookup "on"

___ ENDPROCEDURE lookupcustomers _______________________________________________

___ PROCEDURE forceunlock ______________________________________________________
forceunlockrecord
___ ENDPROCEDURE forceunlock ___________________________________________________

___ PROCEDURE fixphone _________________________________________________________
local newphone
newphone=""
field phone
loop
newphone=phone
newphone=replace(newphone, " ","")
newphone=replace(newphone, "-","")
newphone=replace(newphone, ".","")
newphone=replace(newphone, "+","")
newphone=replace(newphone, "(","")
newphone=replace(newphone, ")","")
phone=newphone[1,3]+"-"+newphone[4,6]+"-"+newphone[7,-1]
;phone=arraychange(phone,"-",4,"")
downrecord
;stop
until info("eof")
newphone=phone
newphone=replace(newphone, " ","")
newphone=replace(newphone, "-","")
newphone=replace(newphone, ".","")
newphone=replace(newphone, "+","")
newphone=replace(newphone, "(","")
newphone=replace(newphone, ")","")
phone=newphone[1,3]+"-"+newphone[4,6]+"-"+newphone[7,-1]
___ ENDPROCEDURE fixphone ______________________________________________________

___ PROCEDURE forcesynchronize _________________________________________________
forcesynchronize
call "listsortcomplete/1"
___ ENDPROCEDURE forcesynchronize ______________________________________________

___ PROCEDURE openall __________________________________________________________
Openfile "customer_hIstory"
Hide
Openfile "45orders"
Openfile "Customer#"
Opensecret "45orders"
Hide

window  waswindow

___ ENDPROCEDURE openall _______________________________________________________

___ PROCEDURE tabdown __________________________________________________________
field Zip
firstrecord
loop
copycell
pastecell
downrecord
stoploopif info("eof")
while forever
copycell
pastecell
___ ENDPROCEDURE tabdown _______________________________________________________

___ PROCEDURE message __________________________________________________________
message str(val(MAd))
___ ENDPROCEDURE message _______________________________________________________

___ PROCEDURE selecttrees ______________________________________________________
serverlookup "off"
select «C#»>0 and «C#»=lookupselected("customer_history", "C#",«C#», "C#",0,0) and T>0 and T<5
serverlookup "on"
___ ENDPROCEDURE selecttrees ___________________________________________________

___ PROCEDURE checkadc's _______________________________________________________
select adc≠lookup("fcmadc","Zip3",val(pattern(Zip,"#####")[1,3]),"adc",0,0) and Zip>0
field adc
formulafill
lookup("fcmadc","Zip3",val(pattern(Zip,"#####")[1,3]),"adc",0,0)
___ ENDPROCEDURE checkadc's ____________________________________________________

___ PROCEDURE deleteunwanted ___________________________________________________
lastrecord
loop
deleterecord
until info("selected")=1
field Con
copy
selectall
find Con=clipboard()
___ ENDPROCEDURE deleteunwanted ________________________________________________

___ PROCEDURE selectduplicates _________________________________________________
field Zip
sortup
field MAd
sortupwithin
selectduplicates MAd+" "+Zip
___ ENDPROCEDURE selectduplicates ______________________________________________

___ PROCEDURE checksize ________________________________________________________
local addressblock
addressblock=""
addressblock=?(Group≠"",Group+¶+Con, Con)+¶+MAd+¶+City+¬+St+¬+pattern(Zip,"#####")
select arraysize(addressblock,¶)>4
___ ENDPROCEDURE checksize _____________________________________________________

___ PROCEDURE newfind __________________________________________________________
case searchcust≠""
liveclairvoyance searchcust, findcust, ¶, "CustomerList","45 mailing list", str(«C#»), "beginswith", str(«C#»), 10, 0, ""
case searchname≠""
liveclairvoyance searchname, findname, ¶, "NameList","45 mailing list", Con, "match", str(«C#»)+": "+Con, 10, 0, ""
endcase
___ ENDPROCEDURE newfind _______________________________________________________

___ PROCEDURE newget ___________________________________________________________
gosheet
find «C#»=val(getcust)
searchcust=""
___ ENDPROCEDURE newget ________________________________________________________

___ PROCEDURE getname __________________________________________________________
gosheet
find Con=extract(getname,": ",2)
searchname=""
call .tab1
___ ENDPROCEDURE getname _______________________________________________________

___ PROCEDURE ExportMail _______________________________________________________
local vCatalogMail
arrayselectedbuild vCatalogMail,¶,"",?(str(«C#»)≠"0", str(«C#»),inqcode)+", "+
Con+", "+Group+", "+MAd+", "+City+", "+St+", "+
?(length(str(Zip))<5,"0"+str(Zip),str(Zip))+", "+«M?»
arraydeduplicate vCatalogMail,vCatalogMail,¶

clipboard()=vCatalogMail
filesave "","CatMail.csv","",vCatalogMail
___ ENDPROCEDURE ExportMail ____________________________________________________

___ PROCEDURE FindBadAddCanada _________________________________________________
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
___ ENDPROCEDURE FindBadAddCanada ______________________________________________

___ PROCEDURE .ExportMacros ____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros from " +info("databasename")+" are saved to your clipboard!"
___ ENDPROCEDURE .ExportMacros _________________________________________________
