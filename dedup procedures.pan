__ PROCEDURE (DeDuplication) __________________________________________________

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

//tallmessage str(nextFYear)+¬+str(thisFYear)+¬+str(lastFYear)


/*

///////~~~~~~~
Programmer Notes
~~~~~~~~~//////////
The danger of this procedure is that come July 1st of the year, it will automatically set
to open the newest files of a non-numbered Panorama file. And if those don't exist, you're 
gonna see errors. Also, a non numbered Panorama file that needs to call older files shouldn't
use this macro



To use these variables please note the following Panorama syntax rules:


filenames using variables:
    can just concatenate as a string
    
    ex:
        
    openfile str(variable)+"filename" 


field calls using variables:
    best to be only one variable and nothing else
    must be surrounded by ( )
    
    ex:
    
    field (VariableFieldName)
    
do your math and/or concatenation into the variable before calling it
    VariableFieldName=str(variable)+"fieldname"
 
field (str(variable)+"fieldname") will work but can cause errors
    
for assignments to that variable'd field 
    use «» for "current field/current cell" 
    
    ex: 
   
    «» = "10"
  
    
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
