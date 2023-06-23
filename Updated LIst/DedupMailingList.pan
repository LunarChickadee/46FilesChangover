___ PROCEDURE .Initialize ______________________________________________________
global New, Num, Numb, enter, place, ID, CC, ED, credit, orders, ship, Flag, waswindow, vfindcurrent, vGlobeSerialNum, DedupOn, mailing_list_window
fileglobal searchcust, findcust, getcust, searchname, findname, getname, vSerialNum
DedupOn=0

mailing_list_window=""


mailing_list_window=info("databasename")

call .AutomaticFY


vSerialNum=""

expressionstacksize 75000000
Num=0
findname=""
searchname=""
searchcust=""
vGlobeSerialNum="0"
gosheet
noshow
call "listsortcomplete/1"
endnoshow
waswindow = info("DatabaseName") 

case lower(folderpath(dbinfo("folder",""))) contains "dedup"
    vSerialNum="2074269900"
    vGlobeSerialNum="2074269900"
    endcase


Openfile "customer_history"

vfindcurrent=0



case folderpath(dbinfo("folder","")) CONTAINS "ogs" or folderpath(dbinfo("folder","")) CONTAINS "walk"
    makesecret
    goform "Add Walkin Customer"
    waswindow = info("windowname")
defaultcase
    openfile "fcmadc"   ///UNCOMMENT THIS
    makesecret
    Openfile "45orders"
    save
    Openfile "46orders"
    save
    openfile "ZipCodeList"
    save
    makesecret
    openfile "Customer#"
    save
    makesecret
    message "All Done Loading Mailing List!"
endcase


window  waswindow
___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE .addmember _______________________________________________________
if info("files") notcontains "members"
openfile "members"
        else
window "members"
        endif
lastrecord
insertbelow
«C#»=grabdata("46 mailing list", «C#»)
Con=grabdata("46 mailing list", Con)
Group=grabdata("46 mailing list", Group)
MAd=grabdata("46 mailing list", MAd)
City=grabdata("46 mailing list", City)
St=grabdata("46 mailing list", St)
Zip=grabdata("46 mailing list", Zip)
SAd=grabdata("46 mailing list", SAd)
Cit=grabdata("46 mailing list", Cit)
Sta=grabdata("46 mailing list", Sta)
Z=grabdata("46 mailing list", Z)
phone=grabdata("46 mailing list", phone)
email=grabdata("46 mailing list", email)
inqcode=grabdata("46 mailing list", inqcode)
«Mem?»=grabdata("46 mailing list", «Mem?»)
windowtoback "members"
window "46 mailing list"
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
global cNumVal,hasAnAddress,hasACon

cNumVal=0
hasACon=""
hasAnAddress=""

field «C#»
    copycell
    cNumVal=val(clipboard())

hasAnAddress=?(MAd≠"",MAd+" "+str(St),"No Mailing Address")

case «Con»="" and «Group»=""
    hasACon = "No Name or Group"
case «Con»≠"" and «Group»≠""
    hasACon = «Con»+"|"+«Group»
defaultcase
    hasACon =«Con»
endcase

YesNo "Delete Customer #:"+str(«C#»)+"?"+¶+"Con: "+hasACon+¶+"MAd: "+hasAnAddress
    if clipboard()="No"
    stop
    endif
deleterecord

if cNumVal>0
        openfile "customer_history"
        openform "customeractivity"
        //___checks if they have information_________________
        case «Con»="" and «Group»=""
            hasACon = "No Name or Group"
        case «Con»≠"" and «Group»≠""
            hasACon = «Con»+"|"+«Group»
        defaultcase
            hasACon =«Con»
        endcase

        hasAnAddress=?(MAd≠"",MAd+" "+str(St),"No Mailing Address")

        //___________________________________________________
        
    if «C#»≠cNumVal
        find «C#»=cNumVal
            if info("found")=0
                window thisFYear+" mailing list"
                stop
            else
                YesNo "delete in customer history?"+¶+str(«C#»)+" "+hasACon+¶+hasAnAddress
                if clipboard()="Yes"
                    deleterecord
                endif
            endif
    endif
endif

window thisFYear+" mailing list"
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
;window thisFYear+" mailing list:customer history"
;else
;downrecord
;endif
if info("formname")="addresschecker"
downrecord
window thisFYear+"orders:seedsinput"
downrecord
window thisFYear+" mailing list:addresschecker"
endif
defaultcase
    key info("modifiers"), KeyStroke
    endcase
___ ENDPROCEDURE .KeyDown ______________________________________________________

___ PROCEDURE .customer ________________________________________________________
OpenFile thisFYear+"orders"
Case Numb>300000 And Numb<400000
    GoForm "ogsinput"
Case Numb>400000 And Numb<500000
    GoForm "treesinput"
Case Numb>500000 And Numb<600000
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
if info("files") contains thisFYear+"orders"
goto continue
else
opensecret thisFYear+"orders"
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
field «Group»
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
Con=grabdata(thisFYear+" mailing list", Con)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
endif
;Con=grabdata(thisFYear+" mailing list", Con)
;«Group»=grabdata(thisFYear+" mailing list", «Group»)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
Notes=replace(Notes, "Bad Address","")
Num=0
endif
window thisFYear+" mailing list"
___ ENDPROCEDURE .MAd __________________________________________________________

___ PROCEDURE .newzip __________________________________________________________
// from editing C# -> calls .custnumber -> if its empty -> calls this

fileglobal listzip, thiszip, findher, findzip, findcity, newcity, findname,findname1, findname2, thisname, firstname, lastname
serverlookup "off" 
;waswindow=info("windowname")
listzip=""
thiszip=""
newcity=""


again:
findher=addressArray


///____gets address from Order user was on or lets user type it in___///
supergettext findher, {caption="Enter Address.Zip" height=100 width=400 captionfont=Courier New captionsize=14 captioncolor="cornflowerblue"
    buttons="Find;Redo;Cancel"}
    if info("dialogtrigger") contains "Find"
        findzip=extract(findher,".",2)
        findzip=strip(findzip)
            if length(findzip)=4
            findzip="0"+findzip
            endif
        findcity=extract(findher,".",1)
        liveclairvoyance findzip,listzip,¶,"","46 mailing list",pattern(Zip,"#####"),"=",str(«C#»)+¬+rep(" ",7-length(str(«C#»)))+Con+rep(" ",max(20-length(Con),1))+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####"),0,0,""
        arraysubset listzip, listzip, ¶, import() contains findcity
            if listzip=""
            goto lastzip
            endif
            
        //___if only one is found, then ask the user if they want to enter it___//
    
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
        liveclairvoyance lastname,findname1,¶,"","46 mailing list",Con,"contains",Con+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")+¬+phone,0,0,""
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
            window "46 mailing list"
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

//not sure if these should be re-enabled
;if info("fieldname")="C#"
;style "record black"
;endif

//_______User Edits C#,Con, or Group____
    //______If they have a customer number, see if they are in customer history already
    //______If not there, create a record for them
Numb=0
if «C#»>0
    Numb=«C#»
    openfile "customer_history"
    Find «C#» = Numb
    if (not info("found"))
        OpenSheet
        lastrecord
        insertbelow
    endif
    «C#»=grabdata(thisFYear+" mailing list", «C#»)
    Con=grabdata(thisFYear+" mailing list", Con)
    «Group»=grabdata(thisFYear+" mailing list", «Group»)
    MAd=grabdata(thisFYear+" mailing list", MAd)
    City=grabdata(thisFYear+" mailing list", City)
    St=grabdata(thisFYear+" mailing list", St)
    Zip=grabdata(thisFYear+" mailing list", Zip)
    «SpareText2»=grabdata(thisFYear+" mailing list", «SpareText2»)
    ;closewindow
    ;window "customer_history:customeractivity"

    window thisFYear+" mailing list"
endif


___ ENDPROCEDURE .tab1 _________________________________________________________

___ PROCEDURE .member __________________________________________________________
if «Mem?»≠"Y"
    stop
endif

if «C#»=0
    Code= thisFYear+"mem"
    openfile "Customer#"
    call "newnumber"

    window thisFYear+" mailing list"
        Field «C#»
        Paste
        SpareText2=str(«C#»)
        if S=0 or T=0 or Bf=0
        call "filler/¬"
endif

field inqcode

window "customer_history"
    insertbelow
    «C#»=grabdata(thisFYear+" mailing list", «C#»)
    «Group»=grabdata(thisFYear+" mailing list", «Group»)
    Con=grabdata(thisFYear+" mailing list", Con)
    MAd=grabdata(thisFYear+" mailing list", MAd)
    City=grabdata(thisFYear+" mailing list", City)
    St=grabdata(thisFYear+" mailing list", St)
    Zip=grabdata(thisFYear+" mailing list", Zip)
    Email=grabdata(thisFYear+" mailing list", email)
    SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
    ;CloseWindow

window thisFYear+" mailing list"
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

window thisFYear+" mailing list"
    call .addmember
___ ENDPROCEDURE .member _______________________________________________________

___ PROCEDURE listsortcomplete/1 _______________________________________________
Hide
//___added this to stop showing mostly empty records on initialize
select str(«C#»)+Con+MAd+City ≠ "0"
//___  -L 8/22
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
local waswindow
waswindow = info("windowname")

if «C#»≠0
    stop 
endif

Code= "I46s"
openfile "Customer#"
call "newnumber"
window waswindow
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
«C#»=grabdata(thisFYear+" mailing list", «C#»)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
Con=grabdata(thisFYear+" mailing list", Con)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
Email=grabdata(thisFYear+" mailing list", email)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
;CloseWindow
window waswindow
___ ENDPROCEDURE createcustomer/µ ______________________________________________

___ PROCEDURE cc rider/ç _______________________________________________________
waswindow=info("windowname")
serverlookup "off"
if info("trigger")="Button.Find Customer #"
if info("files") contains thisFYear+"orders"
goto continue
else
opensecret thisFYear+"orders"
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
OpenFile thisFYear+"orders"
ReSynchronize
field OrderNo
sortup
Case Numb<300000
    GoForm "seedsinput"
Case Numb>300000 And Numb<400000
    GoForm "ogsinput"
Case Numb>400000 And Numb<500000
    GoForm "treesinput"
Case Numb>500000 And Numb<600000
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
global fromBranch

//gets num from record found on mailing list
Num=«C#»


window (thisFYear+"orders")
    case info("formname") contains "ogs" 
        fromBranch = "OGS"
    case info("formname") contains "mt" 
        fromBranch = "OGS;POE"
    case info("formname") contains "bulbs"
        fromBranch = "OGS;Bulbs"
    case info("formanme") contains "Seeds"
        fromBranch = "Seeds"
    case info("formname") contains "Trees"
        fromBranch = "Trees"
    defaultcase 
        fromBranch = ""
    endcase 

window thisFYear + " mailing list"


//___missing number?__//
If Num=0
    call "numberNeeded"
    //adds number to chosen record in ML
    //this also puts them in customer history under their new number
endif

//put that num into Orders
window (thisFYear+"orders")
    «C#»=Num
    «C#Text»=str(Num)
//originally checked if it was an internet order and had an empty name, then stopped

    //call ".IsInternetOrder"

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
Code= "I46s"
openfile "Customer#"
call "newnumber"
window thisFYear+" mailing list"
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
«C#»=grabdata(thisFYear+" mailing list", «C#»)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
Con=grabdata(thisFYear+" mailing list", Con)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
Email=grabdata(thisFYear+" mailing list", email)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
;CloseWindow
window thisFYear+" mailing list"
Call "enter/e"

___ ENDPROCEDURE seedy/ß _______________________________________________________

___ PROCEDURE ogsity/ø _________________________________________________________
if «C#»≠0
stop 
endif
Code="I46o"
openfile "Customer#"
call "newnumber"
window thisFYear+" mailing list"
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
«C#»=grabdata(thisFYear+" mailing list", «C#»)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
Con=grabdata(thisFYear+" mailing list", Con)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
Email=grabdata(thisFYear+" mailing list", email)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
;CloseWindow
window thisFYear+" mailing list"
Call "enter/e"



___ ENDPROCEDURE ogsity/ø ______________________________________________________

___ PROCEDURE moosed/µ _________________________________________________________
if «C#»≠0
stop 
endif
Code= "I46m"
openfile "Customer#"
call "newnumber"
window thisFYear+" mailing list"
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
«C#»=grabdata(thisFYear+" mailing list", «C#»)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
Con=grabdata(thisFYear+" mailing list", Con)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
Email=grabdata(thisFYear+" mailing list", email)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
;CloseWindow
window thisFYear+" mailing list"
Call "enter/e"




___ ENDPROCEDURE moosed/µ ______________________________________________________

___ PROCEDURE treed/† __________________________________________________________
if «C#»≠0
stop 
endif
Code="I46t"
openfile "Customer#"
call "newnumber"
window thisFYear+" mailing list"
Field «C#»
Paste
SpareText2=str(«C#»)
If inqcode=""
Field inqcode
EditCell
field «C#»
EndIf              
call "filler/¬"
endif
if T=0
T=1
endif
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata(thisFYear+" mailing list", «C#»)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
Con=grabdata(thisFYear+" mailing list", Con)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
Email=grabdata(thisFYear+" mailing list", email)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
;CloseWindow
window thisFYear+" mailing list"
Call "enter/e"

___ ENDPROCEDURE treed/† _______________________________________________________

___ PROCEDURE bulbous/∫ ________________________________________________________
if «C#»≠0
stop 
endif
Code= "I46b"
openfile "Customer#"
call "newnumber"
window thisFYear+" mailing list"
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
«C#»=grabdata(thisFYear+" mailing list", «C#»)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
Con=grabdata(thisFYear+" mailing list", Con)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
Email=grabdata(thisFYear+" mailing list", email)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
;CloseWindow
window thisFYear+" mailing list"
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

///called from .newzip and .customerfill
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
        liveclairvoyance findzip,listzip,¶,"",thisFYear+" mailing list",pattern(Zip,"#####"),"=",str(«C#»)+¬+rep(" ",7-length(str(«C#»)))+Con+rep(" ",max(20-length(Con),1))+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####"),0,0,""
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
arraylinebuild address,¬,thisFYear+" mailing list", City+¬+St+¬+str(Zip)+¬+str(adc)+¶
DownRecord
City=extract(address, ¬,1)
St=extract(address, ¬,2)
Zip=val(extract(address, ¬,3))
adc=val(extract(address, ¬,4))
call "filler/¬"
___ ENDPROCEDURE copy city/3 ___________________________________________________

___ PROCEDURE filler/¬ _________________________________________________________
local hasBranchInfo
/* 
added 8/22 by Lunar
*/


window thisFYear+" mailing list"

if S+T+Bf=0 and RedFlag=""
    yesno "- Customer has no catalogs requested"+¶+"- Customer has no RedFlag(s)"+¶+¶+"Autofill catalog requests by Zip/Order?"
    if clipboard()="Yes"
        Case Zip < 19000  And Zip>1000
            S=1
            «M?»=?(«M?» notcontains "X","X"+«M?»,«M?»)
            T=1
            «M?»=?(«M?» notcontains "W","W"+«M?»,«M?»)
            Bf=1
            «M?»=?(«M?» notcontains "Z","Z"+«M?»,«M?»)
        Case (Zip > 43000 And Zip < 46000) 
        or (Zip > 48000 And Zip < 50000) 
        or (Zip > 53000 And Zip < 57000) 
        or Zip>97000
            S=1
            «M?»=?(«M?» notcontains "X","X"+«M?»,«M?»)
            T=1
            «M?»=?(«M?» notcontains "W","W"+«M?»,«M?»)
            Bf=?(fromBranch contains "OGS",1,0)
            «M?»=?(«M?» contains "Z",replace(«M?»,"Z",""),«M?»)
        DefaultCase
            S=1
            «M?»=?(«M?» notcontains "X","X"+«M?»,«M?»)
            T=?(fromBranch contains "Trees",1,0)
            //same for trees and bulbs here
            «M?»=?(«M?» contains "W",replace(«M?»,"W",""),«M?»)
            Bf=?(fromBranch contains "OGS",1,0)
            «M?»=?(«M?» contains "Z",replace(«M?»,"Z",""),«M?»)
        endcase     
    endif 
else 
    case RedFlag≠""
        message "Customer has a RedFlag."+¶+"Catalog requests will be set to zero"
            S=0
            T=0
            Bf=0
            «M?»=""
    defaultcase 
    noyes "Update Catalog Requests?"
    +¶+
    "Currently, Customer is set to receive"
    +¶+
    "Seeds:"+str(S)+" Bulbs:"+str(Bf)+" Trees:"+str(T)
    
    //make this smart enough to only say whaty they're getting?
        if clipboard()="Yes"

        ///this loop is from .UpdateCats
            loop
                rundialog
                “Form="CatalogRequest"
                    Movable=yes
                    okbutton=Update
                    Menus=normal
                    WindowTitle={CatalogRequest}
                    Height=264 Width=190
                    AutoEdit="Text Editor"
                    Variable:"val(«dS»)=val(«S»)"
                    Variable:"val(«dBf»)=val(«Bf»)"
                    Variable:"val(«dT»)=val(«T»)"”
                stoploopif info("trigger")="Dialog.Close"
            while forever 
              message "Customer is now set to receive"
                        +¶+
                        "Seeds:"+str(S)+" Bulbs:"+str(Bf)+" Trees:"+str(T)
                if S≥1 and «M?» notcontains "X"
                    «M?»="X"+«M?»
                else 
                    if S=0
                    «M?»=?(«M?» contains "X",replace(«M?»,"X",""),«M?»)
                    endif
                endif

                if T≥1 and «M?» notcontains "W"
                    «M?»="W"+«M?»
                else 
                    if T=0
                    «M?»=?(«M?» contains "W",replace(«M?»,"W",""),«M?»)
                    endif
                endif

                if Bf≥1 and «M?» notcontains "Z"
                    «M?»="Z"+«M?»
                else 
                    if Bf=0
                    «M?»=?(«M?» contains "Z",replace(«M?»,"Z",""),«M?»)
                    endif
                endif
        endif
    endcase
endif 

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
SelectAdditional «Group» Contains Chr(13)
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
Openfile "customer_history"
Hide
Openfile thisFYear+"orders"
Openfile "Customer#"
Opensecret thisFYear+"orders"
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
addressblock=?(«Group»≠"",«Group»+¶+Con, Con)+¶+MAd+¶+City+¬+St+¬+pattern(Zip,"#####")
select arraysize(addressblock,¶)>4
___ ENDPROCEDURE checksize _____________________________________________________

___ PROCEDURE newfind __________________________________________________________
case searchcust≠""
liveclairvoyance searchcust, findcust, ¶, "CustomerList",thisFYear+" mailing list", str(«C#»), "beginswith", str(«C#»), 10, 0, ""
case searchname≠""
liveclairvoyance searchname, findname, ¶, "NameList",thisFYear+" mailing list", Con, "match", str(«C#»)+": "+Con, 10, 0, ""
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


___ PROCEDURE (DeDuplication) __________________________________________________

___ ENDPROCEDURE (DeDuplication) _______________________________________________

___ PROCEDURE StartDeduplication _______________________________________________
global DedupOn, mailing_list_window

mailing_list_window=""

mailing_list_window=info("databasename")

Define DedupOn,0
//DedupOn=-1

global DedupList,SortThis, con_email_dups,Group_Dups,Phone_Dups,Con_Email_Group_Phone

define con_email_dups,""
define Group_Dups,""
define Phone_Dups,""
define Con_Email_Group_Phone,""

DedupList="Find Same Con&Email [default setup]
Find Same Group Name 
Find Same Phone Number and Con
Find Same Con,Email, Group, and Phone"


superchoicedialog DedupList,SortThis,
{caption="Please choose how to sort for duplicates. 
Please choose NoneOfTheAbove unless you see your name here." captionheight=2 title="User Name Choice"}

;displaydata SortThis

Case SortThis contains "Con&Email"
selectall
Field Con
    sortup
    field email
    sortupwithin
selectduplicates Con+email
selectwithin email ≠ ""
//arrayselectedbuild con_email_dups,¶,"",str(«C#»)+¬+Con
lastrecord


case SortThis contains "Group Name"
selectall
Field Group
    sortup
    field Con
    sortupwithin
selectduplicates Group
selectwithin Group≠""
////arrayselectedbuild Group_Dups,¶,"",str(«C#»)+¬+Con
lastrecord

case SortThis contains "Phone Number"
selectall 
Field phone
    sortup
    field Con
    sortupwithin
selectduplicates phone+Con
selectwithin phone≠""
//arrayselectedbuild Phone_Dups,¶,"",str(«C#»)+¬+Con
lastrecord

case SortThis contains "Con,Email, Group"
selectall 
Field Con
    sortup
    field phone
    sortupwithin
    field email
    sortupwithin
    field Group
    sortupwithin
selectduplicates Con+phone+email+«Group»
selectwithin Con≠"" and phone≠"" and email≠"" 
//arrayselectedbuild Con_Email_Group_Phone,¶,"",str(«C#»)+¬+Con
Field «C#»
lastrecord
endcase


selectwithin «SpareText1» notcontains "duplicate"

lastrecord

___ ENDPROCEDURE StartDeduplication ____________________________________________

___ PROCEDURE SortAndSelect ____________________________________________________

selectall
Field Con
    sortup
    field email
    sortupwithin
selectduplicates Con+email
selectwithin email ≠ ""
lastrecord
___ ENDPROCEDURE SortAndSelect _________________________________________________

___ PROCEDURE SelectSourceRecord/d _____________________________________________
global inq_set,online_set, back_to_ML, ML_Num, dedup_counter_num

define dedup_counter_num, 999999

back_to_ML=info("windowname")
if val(«C#»)=0
    dedup_counter_num=val(dedup_counter_num-1)
    «C#»=val(dedup_counter_num)
endif
copyrecord


window "DeDuplicator"

pasterecord

«C#»=rep(chr(48),6-length(«C#»))+«C#»
LastUpdateSortable=date(Updated[1,"@"][1,-2])
inq_set=?(val(striptonum(inqcode))<30, val(striptonum(inqcode)), val(striptonum(inqcode)[3,4]))
InqCodeNum=inq_set
online_set=val(?(inqcode contains "onl" or inqcode contains "webrequest", "1.2", "1"))
Online=online_set

call "FindMostRecent"

loop
window back_to_ML
until info("windowname") = back_to_ML

___ ENDPROCEDURE SelectSourceRecord/d __________________________________________



___ PROCEDURE .HistoryDelete ___________________________________________________
window "customer_history:customeractivity"
find «C#»=vsourcecust

YesNo "delete this record in customer history?" +" "+ str(«C#») + " "+ Con
if clipboard()="Yes"
deleterecord
endif
window thisFYear+" mailing list"

___ ENDPROCEDURE .HistoryDelete ________________________________________________

___ PROCEDURE .SerialNumber ____________________________________________________


;vSerialNum="" //uncomment this to clear serial numbers
vSerialNum=?(vSerialNum="",info("serialnumber"),vSerialNum+", "+info("serialnumber"))
rudemessage vSerialNum
___ ENDPROCEDURE .SerialNumber _________________________________________________

___ PROCEDURE .Bria's mass delete ______________________________________________
local vCount, vLength
save
saveacopyas "mailinglist backup"
vLength=info("selected")

addrecord

firstrecord

vCount=0

if info("selected")≠info("records")

loop

deleterecord

vCount=vCount+1

until vCount=vLength

selectall

lastrecord

deleterecord

endif
___ ENDPROCEDURE .Bria's mass delete ___________________________________________

___ PROCEDURE (CommonFunctions) ________________________________________________

___ ENDPROCEDURE (CommonFunctions) _____________________________________________

___ PROCEDURE ExportMacros _____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros are saved to your clipboard!"
___ ENDPROCEDURE ExportMacros __________________________________________________

___ PROCEDURE UpdateEmpty ______________________________________________________
select Updated=""
formulafill datepattern(regulardate(Modified),"mm/dd/yy")+"@"+timepattern(regulartime(Modified),"hh:mm am/pm")
___ ENDPROCEDURE UpdateEmpty ___________________________________________________

___ PROCEDURE .UpdateCats ______________________________________________________
            loop
                rundialog
                “Form="CatalogRequest"
                    Movable=yes
                    okbutton=Update
                    Menus=normal
                    WindowTitle={CatalogRequest}
                    Height=264 Width=190
                    AutoEdit="Text Editor"
                    Variable:"val(«dS»)=val(«S»)"
                    Variable:"val(«dBf»)=val(«Bf»)"
                    Variable:"val(«dT»)=val(«T»)"”
                stoploopif info("trigger")="Dialog.Close"
            while forever 
___ ENDPROCEDURE .UpdateCats ___________________________________________________



___ PROCEDURE .appendCustomer __________________________________________________
global appendChoice

appendChoice="default"
appendChoice=str(parameter(1))
if error
appendChoice=appendChoice
endif

////____debug________
;displaydata appendChoice
//__________________

case appendChoice contains "member"
window "members"
lastrecord
insertbelow
«C#»=grabdata("46 mailing list", «C#»)
Con=grabdata("46 mailing list", Con)
Group=grabdata("46 mailing list", Group)
MAd=grabdata("46 mailing list", MAd)
City=grabdata("46 mailing list", City)
St=grabdata("46 mailing list", St)
Zip=grabdata("46 mailing list", Zip)
SAd=grabdata("46 mailing list", SAd)
Cit=grabdata("46 mailing list", Cit)
Sta=grabdata("46 mailing list", Sta)
Z=grabdata("46 mailing list", Z)
phone=grabdata("46 mailing list", phone)
email=grabdata("46 mailing list", email)
inqcode=grabdata("46 mailing list", inqcode)
«Mem?»=grabdata("46 mailing list", «Mem?»)
windowtoback "members"
window thisFYear+" mailing list"

endcase
___ ENDPROCEDURE .appendCustomer _______________________________________________

___ PROCEDURE .hasInfo _________________________________________________________
global cNumVal,hasAnAddress,hasACon

cNumVal=0
hasACon=""
hasAnAddress=""

field «C#»
    copycell
    cNumVal=val(clipboard())

hasAnAddress=?(MAd≠"",MAd+" "+str(Zip),"No Mailing Address")

ç
___ ENDPROCEDURE .hasInfo ______________________________________________________

___ PROCEDURE .TestNewZip ______________________________________________________
// from editing C# -> calls .custnumber -> if its empty -> calls this

fileglobal listzip, thiszip, findher, findzip, findcity, newcity, findname,findname1, findname2, thisname, firstname, lastname
serverlookup "off" 
;waswindow=info("windowname")
listzip=""
thiszip=""
newcity=""


again:
findher=addressArray


///____gets address from Order user was on or lets user type it in___///
supergettext findher, {caption="Enter Address.Zip" height=100 width=400 captionfont=Courier New captionsize=14 captioncolor="cornflowerblue"
    buttons="Find;Redo;Cancel"}
    if info("dialogtrigger") contains "Find"
        findzip=extract(findher,".",2)
        findzip=strip(findzip)
            if length(findzip)=4
            findzip="0"+findzip
            endif
        findcity=extract(findher,".",1)
        liveclairvoyance findzip,listzip,¶,"","46 mailing list",pattern(Zip,"#####"),"=",str(«C#»)+¬+rep(" ",7-length(str(«C#»)))+Con+rep(" ",max(20-length(Con),1))+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####"),0,0,""
        arraysubset listzip, listzip, ¶, import() contains findcity
            if listzip=""
            goto lastzip
            endif
            
        //___if only one is found, then ask the user if they want to enter it___//
    
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
        liveclairvoyance lastname,findname1,¶,"","46 mailing list",Con,"contains",Con+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")+¬+phone,0,0,""
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
            window "46 mailing list"
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

___ ENDPROCEDURE .TestNewZip ___________________________________________________

___ PROCEDURE vGetName _________________________________________________________
gosheet
find Con=extract(getname,": ",2)
searchname=""
call .tab1
___ ENDPROCEDURE vGetName ______________________________________________________

___ PROCEDURE numberNeeded _____________________________________________________
//if you got here, but they have a number, skip this whole process and go back to enter
if «C#»≠0 
    goto End
endif

//gives them an First Year and Branch of Order code
    case fromBranch contains ";"
        if fromBranch contains "Bulb"
            Code="I"+thisFYear+"b"
        endif
        if fromBranch contains "POE"
            Code="I"+thisFYear+"p"
        endif
    case fromBranch contains "Seeds"
    Code = "I"+thisFYear+"s"
    case fromBranch contains "OGS"
    Code = "I"+thisFYear+"o"
    case fromBranch contains "Trees"
    Code = "I"+thisFYear+"t"
    endcase

//gets a new number
openfile "Customer#"
call "newnumber"
window thisFYear+" mailing list"
Field «C#»
    Paste

//saves the C# in sparetext2
SpareText2=str(«C#»)

//Get InqCode
If inqcode=""

Field inqcode
    //autofill internet entry
    case internet_order=-1
        inqcode=str(yearvalue(today()))[3,4]+"onl"

    defaultcase
        editcell
        //if it has this year twice e.g. "2222", then just make it the 22 then inqcode
        inqcode=?(inqcode contains str(yearvalue(today()))[3,4]+str(yearvalue(today()))[3,4],inqcode[3,-1],inqcode)
        //if someone types 2022 goes to 7th character because it'll autofill as 222022onl
        inqcode=?(inqcode contains str(yearvalue(today())),str(yearvalue(today()))[3,4]+inqcode[7,-1],inqcode)
    endcase
field «C#»
EndIf

//fills in catalog requests based on order/location/redflags
call "filler/¬"

If inqcode=""
field inqcode
editcell
endif


//_______give them customer history________________________//
window "customer_history:secret"
opensheet
insertbelow
«C#»=grabdata(thisFYear+" mailing list", «C#»)
«Group»=grabdata(thisFYear+" mailing list", «Group»)
Con=grabdata(thisFYear+" mailing list", Con)
MAd=grabdata(thisFYear+" mailing list", MAd)
City=grabdata(thisFYear+" mailing list", City)
St=grabdata(thisFYear+" mailing list", St)
Zip=grabdata(thisFYear+" mailing list", Zip)
Email=grabdata(thisFYear+" mailing list", email)
SpareText2=grabdata(thisFYear+" mailing list", SpareText2)
;CloseWindow

End:
window thisFYear+" mailing list"
Call "enter/e"

___ ENDPROCEDURE numberNeeded __________________________________________________

___ PROCEDURE .OpenOrderSearch/¬ _______________________________________________
window thisFYear+"orders"
call NewSearch/¬
___ ENDPROCEDURE .OpenOrderSearch/¬ ____________________________________________

___ PROCEDURE .getequations ____________________________________________________
global eqArray
arraybuild eqArray, ¶,"", ?(Equation≠"","~~"+«Field Name»+"~~"+¶+Equation+¶,"")
clipboard()=eqArray 
___ ENDPROCEDURE .getequations _________________________________________________

___ PROCEDURE .test ____________________________________________________________
message info("datatype")
___ ENDPROCEDURE .test _________________________________________________________

___ PROCEDURE .test3 ___________________________________________________________
global Con_Email_Group_Phone
selectall 
Field Con
    sortup
    field phone
    sortupwithin
    field email
    sortupwithin
    field Group
    sortupwithin
selectduplicates Con+phone+email+«Group»
;selectwithin Con≠"" and phone≠"" and email≠"" 
arrayselectedbuild Con_Email_Group_Phone,¶,"",str(«C#»)+¬+Con
Field «C#»
lastrecord


___ ENDPROCEDURE .test3 ________________________________________________________

___ PROCEDURE .AutomaticFY _____________________________________________________
global dateHold, dateMath, intYear, 
thisFYear,lastFYear,nextFYear,intMonth,fileDate

thisFYear = "46"
lastFYear = "45"
nextFYear = "47"

//______this does the FY automatically, but has issues if people want to open multiple instances

/*
fileDate=val(striptonum(info("databasename")))
nextFYear=""
thisFYear=""
lastFYear=""

//get the date
dateHold = datepattern(today(),"mm/yyyy")

//gets the current month and year
intMonth = val(dateHold[1,"/"][1,-2])
intYear = val(dateHold["/",-1][2,-1])

//assigns FY numbers for years

case val(intMonth)>6
    nextFYear=str(intYear-1976)
    thisFYear=str(intYear-1977)
    lastFYear=str(intYear-1978)

case val(intMonth)<7
    nextFYear=str(intYear-1977)
    thisFYear=str(intYear-1978)
    lastFYear=str(intYear-1979)

endcase

//checks if this is an older file and needs older FYs
if fileDate ≤ val(lastFYear) and fileDate > 0
    message "This file is older than 2 years and will try to open older order files, etc"
    nextFYear=str(fileDate+1)
    thisFYear=str(fileDate)
    lastFYear=str(fileDate-1)
endif

*/


___ ENDPROCEDURE .AutomaticFY __________________________________________________

___ PROCEDURE .SetDedupCounter _________________________________________________

___ ENDPROCEDURE .SetDedupCounter ______________________________________________

___ PROCEDURE SelectionLoop ____________________________________________________
noshow

global list_count

list_count = ""

firstrecord

loop

list_count = str(«C#»)+","+list_count

downrecord

until 20000

arraystrip list_count, ","

endnoshow

select list_count contains str(«C#»)
___ ENDPROCEDURE SelectionLoop _________________________________________________
