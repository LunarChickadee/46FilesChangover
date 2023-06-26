___ PROCEDURE .ZeroOutCatalogs _________________________________________________
local vS,vT,vBf

vS=S
vT=T
vBf=Bf

if RedFlag contains "No Catalogs"
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

___ PROCEDURE listsortcomplete/1 _______________________________________________
hide
SelectAll
save
noUndo
InsertBelow
Field "MAd"
SortUp
Field "City"
SortUp
Field "St"
SortUp
Field "Zip"
SortUp

Select «C#»≠0
SelectAdditional St=""
SelectAdditional inqcode BEGINSWITH str(20)
SelectReverse
Field inqcode
FormulaFill str(20)+inqcode
SelectAll
Field Con
FirstRecord
show
___ ENDPROCEDURE listsortcomplete/1 ____________________________________________

___ PROCEDURE printMulti/2 _____________________________________________________
global zselect
AddRecord
Select val(«M?»["0-9",-1])>0
if info("selected")=info("records")
    message "no multiples today, go directly to macro 3"
    zselect =  info("Selected") 
    stop
endif
Message "Put a check stub in the printer--don't need labels for these!"
OpenForm "MultiSpecial"
Print ""
CloseWindow
GoForm "inquiryentryView"
SelectReverse
Save
zselect =  info("Selected") 
Message "if print job ran OK, go to 3 macro to print Seeds labels"

___ ENDPROCEDURE printMulti/2 __________________________________________________

___ PROCEDURE printSeeds/3 _____________________________________________________
save
zselect =  info("Selected") 
SelectWithin «M?» contains "X"
if info("selected")=zselect
    Message "No Seed catalogs today — go on to macro 4"
    stop
endif
Message "Make sure there are labels in the printer !"
OpenForm "catalog label"
   zoomwindow 285,800,150,400,""
Print ""
GoForm "inquiryentryView"
Revert
Message "if labels printed OK, go to 4 macro to print trees"

___ ENDPROCEDURE printSeeds/3 __________________________________________________

___ PROCEDURE printTrees/4 _____________________________________________________
SelectWithin «M?» contains "W"
if info("selected")=info("records")
    Message "No Tree catalogs today — go on to macro 5"
    GoForm "catalog label"
    CloseWindow
    GoForm "entrybox"
    stop
endif
Message "Make sure there are labels in the printer !!"
GoForm "catalog label"
Print ""
CloseWindow
GoForm "inquiryentryView"
GoForm "entrybox"
SelectAll
save
Message "if everything's OK, go to 5 macro to finish up"

___ ENDPROCEDURE printTrees/4 __________________________________________________

___ PROCEDURE finalize/5 _______________________________________________________
local check, inqwindow, vEmail, vName, vOnlist, vAdded, vCount1, vLength
inqwindow=info("windowname")

    check=""
    vEmail=""
;******** Fills Date field **********    
field «date mailing»
    GetScrapOK "date of mailing?"
    Fill clipboard()
;*******************************
;hide
;noshow
selectall
firstrecord
;********* Builds Arrays for searching for matches******
openfile "46 mailing list"
    selectall
    arraybuild vEmail,¶,"",email
    arraybuild vName,¶,"",replace(Con," ","")+?(length(str(Zip))<5,"0"+str(Zip),str(Zip))
    arraystrip vName,¶

        clipboard()=vName

;****************************************************

;********* Searches inqlist file and fills "onlist" field *******
window inqwindow
    field onlist
    formulafill ?(vEmail contains email, "onlist",
                ?(vName contains replace(Con," ","")+
                ?(length(str(Zip))<5,"0"+str(Zip),str(Zip)),"onlist",""))
                     
                     ;stop
                    
;****************************************************

;endnoshow
;show

;********* Appends "InquiryArchive" file **********
OpenFile "inquiryArchive"
    OpenFile "++Inquiry List"
    save
CloseFile
;*********************************************

field «M?»
    Fill ""
save

;*********** Builds a selection to add to the mailing list ***************
;select Con≠""
    ;RemoveUnselected
    
//replaced by Rachel to work in a linked file


select MAd=""
vLength=info("selected")
addrecord
firstrecord
vCount1=0
if info("selected")≠info("records")
    loop
        deleterecord
        vCount1=vCount1+1
    until vCount1=vLength
    selectall
    lastrecord
    deleterecord
endif


select onlist=""
arraybuild vOnlist,¶,"",onlist
;message vOnlist


YesNo "Do you want to append to mailing list?"
    If clipboard()="Yes"
        if info("selected")=info("records") and vOnlist contains "onlist"
           
            message "All records are on the mailing list, nothing to append"
           
            goto Inquiry
        endif
        vAdded=info("selected")
        message "Adding "+str(vAdded)+" records to the mailing list."
        window "46 mailing list"
        ;SelectAll
        ;message info("files")
        if  info("Files") notcontains "46 mailing list"
            message "Records not appended. Problem opening mailing list. Do not rerun this macro. Get help."
            stop
        endif
          
        openfile "++Inquiry List"
        save
        Window "Hide This Window"
    endif
;*****************************************************************
  Inquiry:
window inqwindow
SelectAll

addrecord
firstrecord
loop
    deleterecord
until info("records")=1


//Previous Delete method when not linked
;deleteall


save

___ ENDPROCEDURE finalize/5 ____________________________________________________

___ PROCEDURE PrintBulbs _______________________________________________________
SelectWithin «M?» contains "Z"
if info("selected")=zselect
    stop
endif
Message "Make sure there are labels in the printer !!"
OpenForm "catalog label"
Print ""
GoForm "inquiryentryView"
SelectAll
save

___ ENDPROCEDURE PrintBulbs ____________________________________________________

___ PROCEDURE .windowtoback ____________________________________________________
WindowToBack "Inquiry List:inquiryentry"
___ ENDPROCEDURE .windowtoback _________________________________________________

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

if «M?» contains "W" and T=0
    T=1
endif

if «M?» contains "Z" and Bf=0
    Bf=1
endif

if «M?» contains "X" and S=0
    S=1
endif
save

call .ZeroOutCatalogs

___ ENDPROCEDURE filler/¬ ______________________________________________________

___ PROCEDURE goofy ____________________________________________________________
OpenFile "34 mailing list"

___ ENDPROCEDURE goofy _________________________________________________________

___ PROCEDURE zipper/Ω _________________________________________________________
getscrap "What's the zip?"
Num=val(clipboard())
Find Zip=Num

___ ENDPROCEDURE zipper/Ω ______________________________________________________

___ PROCEDURE callZip/ª ________________________________________________________
global jcity, jstate, jzip, jtowns
jzip=0
GetScrap "What's the Zip Code?"
if val(clipboard())<601 or val(clipboard())>99999
    Message "That is a BOOOOGUS Zip code. Try again."
    stop
endif
LastRecord
InsertBelow
;Zip=val(clipboard())
field Zip
Paste


OpenFile "ycitylist"
SelectAll
Field Zip
SortUp
save
Find Zip=val(clipboard())
jzip=val(clipboard())

if info("found")
    jcity=City
    jstate=St

    YesNo  "Is "+«City»+", "+«St»+"  right ?"
    ;if typed correctly
        if clipboard()="Yes"
            window "Inquiry List:entrybox"
            LastRecord
            Zip=jzip
            St=jstate
            City=jcity
            field Con
        endif
    ;if town not right
        if clipboard()="No"
            YesNo "You sure  "+pattern(jzip,"#####")+"  is OK?"
        
            if clipboard()="Yes"    
                arraybuild jtowns,¶,"",?(Zip=jzip,City,"")
                if arraysize(jtowns,¶)=1
                    Message "Sorry, that's all the towns we have today. Please check the USPS website for all towns with this zip."
                    window "Inquiry List:entrybox"
                    stop
                endif
            local jnum
            jnum=1
            loop
                jnum=jnum+1
                jcity=extract(jtowns,¶,jnum)
                YesNo  "Is "+jcity+", "+jstate+"  right ?"
            until clipboard()="Yes" or jnum=arraysize(jtowns,¶)
                if clipboard()="No"
                    Message "Sorry, that's all the towns we have today. Please check the USPS website for all towns with this zip."
                    window "Inquiry List:entrybox"
                    stop
                endif
                if clipboard()="Yes"
                    window "Inquiry List:entrybox"
                    LastRecord
                    Zip=jzip
                    St=jstate
                    City=jcity
                    field Con
                endif
            endif
        if clipboard()="No"
            window "Inquiry List:entrybox"
            call callZip/ª
        endif
    endif

    else
        YesNo "You sure  "+pattern(jzip,"#####")+"  is OK?"
        if clipboard()="Yes"
            BigMessage "That Zip is not in our database. "+¶+"• Check the Post Office website to confirm the Zip and the spelling of the town.  "+¶+"• Then go to MelissaData and confirm the street. "+¶+¶+"• Then use the New Zip Code button."
        endif
        if clipboard()="No"
            window "Inquiry List:entrybox"
            call callZip/ª
        endif
    endif
    window "Inquiry List:entrybox"
    Field Con
    save
    
___ ENDPROCEDURE callZip/ª _____________________________________________________

___ PROCEDURE getCon ___________________________________________________________
if Con=""
    field Group
    Cut
    field Con
    Paste
    field inqcode
endif

___ ENDPROCEDURE getCon ________________________________________________________

___ PROCEDURE codeCount ________________________________________________________
call .Inquiry

local vDate, vSeed, vTree, vBulb
vSeed="" 
vTree="" 
vBulb=""

if «M?» contains "X"
    vSeed="S"
else 
    vSeed=""
endif


if «M?» contains "W"
    vTree="T"
else
    vTree=""
endif


if «M?» contains "Z"
    vBulb="B"
else
    vBulb=""
endif

vDate=datepattern(today(),"YY")
inqcode=?(inqcode="","",vDate+vCode+" - "+vSeed+vTree+vBulb)





call filler/¬

___ ENDPROCEDURE codeCount _____________________________________________________

___ PROCEDURE newZipCode/º _____________________________________________________
global jcity, jstate, jzip, jtowns
jzip=Zip
jcity=0
jstate=0
YesNo  "You going with "+pattern(jzip,"#####")+",  right ?"
    if clipboard()="Yes"
        jzip=Zip
    endif

    if clipboard()="No"
        ;field Zip
        ;Zip=0
        GetScrap "What is the new Zip Code?"
        ;Zip=val(clipboard())
        jzip=Zip
    endif

   ;stop
OpenFile "ycitylist"
InsertRecord
Zip=jzip
OpenForm "NewTown"
zoomwindow 483,800,400,600,""
Field City
EditCell
CloseWindow
Field Zip
SortUp
UpRecord
Field St
copy
DownRecord
paste
Save
jzip=Zip
jcity=City
jstate=St
Window "Hide This Window"
window "Inquiry List:entrybox"
if info("records")=info("selected")
    LastRecord
    Zip=jzip
    City=jcity
    St=jstate
    Field Con
endif
;lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"St","",0)

;stop
;lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)


___ ENDPROCEDURE newZipCode/º __________________________________________________

___ PROCEDURE hide window ______________________________________________________
Window "Hide This Window"

___ ENDPROCEDURE hide window ___________________________________________________

___ PROCEDURE canadiens ________________________________________________________
LastRecord
InsertBelow
OpenForm "provinces"
zoomwindow 483,800,400,400,""
___ ENDPROCEDURE canadiens _____________________________________________________

___ PROCEDURE finishCanadiens __________________________________________________
St= info("Trigger")[".",-1][2,-1]
CloseWindow
;get postal code
OpenForm "postal code"
zoomwindow 483,800,400,500,""
Field Sta
EditCell
CutCell
CloseWindow
Field St
    if clipboard()[4,4]≠ " "
    clipboard()=clipboard()[1,3]+" "+clipboard()[4,6]
    endif
St=St+" "+clipboard()
;get town
OpenForm "NewTown"
zoomwindow 483,800,400,600,""
Field City
EditCell
CloseWindow
field Con





___ ENDPROCEDURE finishCanadiens _______________________________________________

___ PROCEDURE info _____________________________________________________________
message info("files")
___ ENDPROCEDURE info __________________________________________________________

___ PROCEDURE .Initialize ______________________________________________________
global vBranch,vContact,vConference,vInqCode,vSelect,vSelect2,vSelect3,vYear,vList,vHowHeard
vSelect=""
vSelect2=""
vSelect3=""
vList=""
vYear=str(yearvalue(today()))[3,4]
;*****Use these lists to build INQ Menus******
;*****Seperate each item with a comma*****
vBranch="Bulbs,Conference,OGS,Seeds,Trees"
vContact="Email,SnailMail,WalkIn"
vConference="MOSES,NOFA,CGCF"
vHowHeard="Friend,OGOrg,Book,ExtentionService,Internet,SeedsCompany,Referral,Paper,Magazine,ExistingCustomer"

;****************************************
;openSecret "newadc"
openfile "ycitylist"
Window "ycitylist"
Window "Hide This Window"
window "Inquiry List"
;openform "inquiryentryView"


___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE tagtown/• ________________________________________________________
global jcity, jstate, jzip, jtowns
jzip=0
GetScrap "What's the Zip Code?"
if val(clipboard())<601 or val(clipboard())>99999
    Message "That is a BOOOOGUS Zip code. Try again."
stop
endif
;Zip=val(clipboard())
field Zip
Paste


OpenFile "ycitylist"
SelectAll
Field Zip
SortUp
save
Find Zip=val(clipboard())
jzip=val(clipboard())

if info("found")
jcity=City
jstate=St

YesNo  "Is "+«City»+", "+«St»+"  right ?"
;if typed correctly
    if clipboard()="Yes"
        window "Inquiry List:entrybox"
        LastRecord
        Zip=jzip
        St=jstate
        City=jcity
        field Con
    endif
;if town not right
    if clipboard()="No"
        YesNo "You sure  "+pattern(jzip,"#####")+"  is OK?"
    
            if clipboard()="Yes"    
                arraybuild jtowns,¶,"",?(Zip=jzip,City,"")
                if arraysize(jtowns,¶)=1
                    Message "Sorry, that's all the towns we have today. Please check the USPS website for all towns with this zip."
                    window "Inquiry List:entrybox"
                    stop
                endif
            local jnum
            jnum=1
            loop
                jnum=jnum+1
                jcity=extract(jtowns,¶,jnum)
                YesNo  "Is "+jcity+", "+jstate+"  right ?"
            until clipboard()="Yes" or jnum=arraysize(jtowns,¶)
            if clipboard()="No"
                Message "Sorry, that's all the towns we have today. Please check the USPS website for all towns with this zip."
                window "Inquiry List:entrybox"
                stop
            endif
            if clipboard()="Yes"
                window "Inquiry List:entrybox"
                LastRecord
                Zip=jzip
                St=jstate
                City=jcity
                field Con
            endif
        endif
        if clipboard()="No"
            window "Inquiry List:entrybox"
            call tagtown/•
        endif
    endif

    else
        YesNo "You sure  "+pattern(jzip,"#####")+"  is OK?"
        if clipboard()="Yes"
            BigMessage "That Zip is not in our database. "+¶+"• Check the Post Office website to confirm the Zip and the spelling of the town.  "+¶+"• Then go to MelissaData and confirm the street. "+¶+¶+"• Then use the New Zip Code button."
        endif
        if clipboard()="No"
            window "Inquiry List:entrybox"
            call tagtown/•
        endif
    endif
window "Inquiry List:entrybox"
Field inqcode
save

___ ENDPROCEDURE tagtown/• _____________________________________________________

___ PROCEDURE addrecord/7 ______________________________________________________
;inserts blank line to end
AddRecord
Field Con
___ ENDPROCEDURE addrecord/7 ___________________________________________________

___ PROCEDURE count/ç __________________________________________________________
Field "M?"
Select «M?» contains "X"
message "This many Seed cats: "+str(info("selected"))
SelectAll
Select «M?» contains "W"
message "This many Trees cats: "+str(info("selected"))
SelectAll
LastRecord
___ ENDPROCEDURE count/ç _______________________________________________________

___ PROCEDURE huh ______________________________________________________________
local check, inqwindow, vEmail, vName
inqwindow=info("windowname")

    check=""
    vEmail=""
;******** Fills Date field **********    
field «date mailing»
    GetScrapOK "date of mailing?"
    Fill clipboard()
;*******************************
;hide
;noshow
selectall
firstrecord
;********* Builds Arrays for searching for matches******
window "42 mailing list"
    selectall
    arraybuild vEmail,¶,"",email
    arraybuild vName,¶,"",replace(Con," ","")+?(length(str(Zip))<5,"0"+str(Zip),str(Zip))
    arraystrip vName,¶

        clipboard()=vName

;****************************************************

;********* Searches inqlist file and fills "onlist" field *******
window inqwindow
    field onlist
    formulafill ?(vEmail contains email, "onlist",
                     ?(vName contains replace(Con," ","")+
                     ?(length(str(Zip))<5,"0"+str(Zip),str(Zip)),"onlist",""))
                    
;****************************************************

;endnoshow
;show

;********* Appends "InquiryArchive" file **********
OpenFile "inquiryArchive"
    OpenFile "++Inquiry List"
    save
CloseFile
;*********************************************

field «M?»
    Fill ""
save

;*********** Builds a selection to add to the mailing list ***************
select Con≠""
    RemoveUnselected
select onlist=""

YesNo "Do you want to append to mailing list?"
    If clipboard()="Yes"
        if info("selected")=info("records")
            message "All records are on the mailing list, nothing to append"
            goto Inquiry
        endif
        
        OpenFile "42 mailing list"
        ;SelectAll
        ;message info("files")
        if  info("Files") notcontains "42 mailing list"
            message "Records not appended. Problem opening mailing list. Do not rerun this macro. Get help."
            stop
        endif
        Inquiry:
        openfile "++Inquiry List"
        save
        CloseFile  "42 mailing list"
        Window "Hide This Window"
    endif
;*****************************************************************
window inqwindow
SelectAll

deleteall
save

___ ENDPROCEDURE huh ___________________________________________________________

___ PROCEDURE (Import Requests) ________________________________________________

___ ENDPROCEDURE (Import Requests) _____________________________________________

___ PROCEDURE Import Web Requests ______________________________________________
;**********************************************************
;******This routine imports catalog requests downloaded from the website
;****** scrubbs it for special characters then checks the address in the zipcode list
;**********************************************************

local vFolder,vFile,vType

NoYes "Do you want to replace your current data?"

if clipboard()="Yes"
    openfiledialog vFolder,vFile,vType,""; "&catalog_requests.csv"
    opentextfile "&"+folderpath(vFolder)+vFile
    else
        NoYes "Would you like to append your data to this file?"
        if clipboard()="Yes"
            openfiledialog vFolder,vFile,vType,""; "&catalog_requests.csv"
            opentextfile "+"+folderpath(vFolder)+vFile
            else
                message "Data was not imported"
                stop
        endif
endif


 
 call .CleanAddresses
     
call .CheckAddress

call codeCount

field Con


___ ENDPROCEDURE Import Web Requests ___________________________________________

___ PROCEDURE .CheckAddress ____________________________________________________
;**********************************************************
;***** This routine builds a selection of addresses that need some attention before mailing
;**********************************************************


global vAddress,vAddress2
local vCanada, vRecCount
vRecCount=info("records")

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

;select lookup("ycitylist","Zip",«Zip»,"Zip",0,0)
firstrecord
arraybuild vAddress,¶,"ycitylist",str(Zip)+St+City
clipboard=vAddress
vAddress2=""
loop
    clipboard()=vAddress
   ; rudemessage vAddress
;    stop
    if vAddress notcontains  str(Zip)+St 
        vAddress2=vAddress2+¶+str(Zip)+St+City   
    endif
    downrecord
until info("stopped")
; rudemessage vAddress2
select vAddress2 notcontains str(Zip)+St
;stop
selectadditional length(St)>2

selectreverse
selectadditional City contains "@"
Selectadditional MAd contains "@"
selectadditional vCanada contains St
if info("records")≠info("selected")
    bigmessage "Something is wrong with these adreeses. Please double check them." +¶+¶+¶+"If you see Canadian addressed here, they need their full Postal Code in the «St» Field"
endif


___ ENDPROCEDURE .CheckAddress _________________________________________________

___ PROCEDURE .CleanAddresses __________________________________________________
//2022 version Rachel M


;**********************************************************
;***** This routine cleansup fields to make them more suitable for the mail.
;**********************************************************


field Con
    serverformulafill str(val(Con)) notcontains "0", ?(Con contains "@", "Resident", Con)
    
    
field Group
    serverformulafill str(val(Group)) notcontains "0", replace(Group, "
","")
    serverformulafill str(val(Group)) notcontains "0", replace(Group, ".","")
    serverformulafill str(val(Group)) notcontains "0", replace(Group, ","," ")
    serverformulafill str(val(Group)) notcontains "0", replace(Group,"‚Äô","'")
    serverformulafill str(val(Group)) notcontains "0", upperword(Group)
field Con
    serverformulafill str(val(Con)) notcontains "0", replace(Con, "
","")
    serverformulafill str(val(Con)) notcontains "0", replace(Con, ".","")
    serverformulafill str(val(Con)) notcontains "0", replace(Con, ","," ")
    serverformulafill str(val(Con)) notcontains "0", replace(Con,"‚Äô","'")
    serverformulafill str(val(Con)) notcontains "0", upperword(Con)
field MAd
    serverformulafill str(val(MAd)) notcontains "0", replace(MAd, "
","")
    serverformulafill str(val(MAd)) notcontains "0", replace(MAd, ".","")
    serverformulafill str(val(MAd)) notcontains "0", replace(MAd, ","," ")
    serverformulafill str(val(MAd)) notcontains "0", replace(MAd,"---","")
    serverformulafill str(val(MAd)) notcontains "0", replace(MAd,"‚Äô","'")
    serverformulafill str(val(MAd)) notcontains "0", upperword(MAd)
field City
   serverformulafill str(val(City)) notcontains "0", replace(City, "
","")
    serverformulafill str(val(City)) notcontains "0", replace(City, ".","")
    serverformulafill str(val(City)) notcontains "0", replace(City, ","," ")
    serverformulafill str(val(City)) notcontains "0", replace(City,"‚Äô","'")
    serverformulafill str(val(City)) notcontains "0", upperword(City)
field Zip
    serverformulafill str(val(Zip)) notcontains "@", val(str(Zip)[1,5])
    serverformulafill str(val(Zip)) notcontains "@", zeroblank(Zip)







/* Previous Version 2021
;**********************************************************
;***** This routine cleansup fields to make them more suitable for the mail.
;**********************************************************


field Con
    formulafill ?(Con contains "@", "Resident", Con)
field Group
    formulafill replace(Group, "
","")
    formulafill replace(Group, ".","")
    formulafill replace(Group, ","," ")
    formulafill upperword(Group)
field Con
    formulafill replace(Con, "
","")
    formulafill replace(Con, ".","")
    formulafill replace(Con, ","," ")
    formulafill upperword(Con)
field MAd
    formulafill replace(MAd, "
","")
    formulafill replace(MAd, ".","")
    formulafill replace(MAd, ","," ")
    formulafill replace(MAd,"---","")
    formulafill upperword(MAd)
field City
    formulafill replace(City, "
","")
    formulafill replace(City, ".","")
    formulafill replace(City, ","," ")
    formulafill upperword(City)
field Zip
    formulafill val(str(Zip)[1,5])
    formulafill zeroblank(Zip)
    
    */
___ ENDPROCEDURE .CleanAddresses _______________________________________________

___ PROCEDURE .BuildINQ ________________________________________________________
vSelect2=""
vSelect3=""

;message vInqCode
vList=?(vSelect="Bulbs",vContact,
        ?(vSelect="Seeds",vContact,
        ?(vSelect="OGS",vContact,
        ?(vSelect="Trees",vContact,vConference))))

popupatmouse replace(vList,",",¶),"",vSelect2
if vSelect≠"Conference"
    popupatmouse replace(vHowHeard,",",¶),"",vSelect3
endif

drawobjects
vInqCode=str(vYear)+"-"+vSelect+"-"+vSelect2+?(vSelect3="","","-"+vSelect3)
bigmessage vInqCode
inqcode=vInqCode

___ ENDPROCEDURE .BuildINQ _____________________________________________________

___ PROCEDURE huh2 _____________________________________________________________
local firstname, lastname, lzip, check, inqwindow, vEmail
inqwindow=info("windowname")

field «date mailing»
    GetScrapOK "date of mailing?"
    Fill clipboard()

;hide
;noshow
selectall
firstrecord
loop
    firstname=""
    lastname=""
    check=""

    firstname=Con[1," "][1,-2]
    lastname=Con[" ",-1][2,-1]
    lzip=Zip
    
    window "42 mailing list"
    selectall
    arraybuild vEmail,¶,"",email
    find Con contains firstname and Con contains lastname
    if info("found")=0
        check=""
    endif
    if info("found")<0
        find  Con contains firstname and Con contains lastname and Zip=lzip
            if info("found")=-1
                check="onlist"
           ;else
               ;check="maybe"
            endif
    endif
    
    window inqwindow
        onlist=check
    downrecord
until info("stopped")

field onlist
    formulafill ?(vEmail contains email, "onlist",onlist)

;endnoshow
;show


OpenFile "inquiryArchive"
    OpenFile "++Inquiry List"
    save
CloseFile

field «M?»
    Fill ""
save

select Con≠""
    RemoveUnselected
select onlist=""
YesNo "Do you want to append to mailing list?"
    If clipboard()="Yes"
        OpenFile "42 mailing list"
        ;SelectAll
        ;message info("files")
            if  info("Files") notcontains "42 mailing list"
                message "Records not appended. Problem opening mailing list. Do not rerun this macro. Get help."
                stop
            endif
        openfile "++Inquiry List"
        save
        ;CloseFile  "42 mailing list"
        Window "Hide This Window"
    endif

window inqwindow
SelectAll

;deleteall
save

___ ENDPROCEDURE huh2 __________________________________________________________

___ PROCEDURE .ClearMBeta ______________________________________________________
local vLet
vLet=«M?»
getscrap "Which Letter Would you Like to clear?"
field «M?»


if clipboard()="x"
    vLet=stripchar(vLet,"WZ")
    «M?»=vLet
endif

___ ENDPROCEDURE .ClearMBeta ___________________________________________________

___ PROCEDURE ExportMail _______________________________________________________
local vCatalogMail
arrayselectedbuild vCatalogMail,¶,"",?(str(«C#»)≠"0", str(«C#»),inqcode)+", "+
                                    Con+", "+Group+", "+MAd+", "+City+", "+St+", "+
                                    ?(length(str(Zip))<5,"0"+str(Zip),str(Zip))+", "+«M?»
arraydeduplicate vCatalogMail,vCatalogMail,¶
                                                
clipboard()=vCatalogMail
filesave "","CatMail.csv","",vCatalogMail


___ ENDPROCEDURE ExportMail ____________________________________________________

___ PROCEDURE .Inquiry _________________________________________________________
global vCode
vCode=inqcode

if vCode contains "Phone"
    vCode="Phone"
endif
if vCode contains "Email"
    vCode="Email"
endif
if vCode contains "webrequest"
    vCode="webrequest"
endif




___ ENDPROCEDURE .Inquiry ______________________________________________________

___ PROCEDURE DeleteTest _______________________________________________________

addrecord
firstrecord
loop
    deleterecord
until info("records")=1



___ ENDPROCEDURE DeleteTest ____________________________________________________

___ PROCEDURE .NoCatalogsMacros ________________________________________________
///unfinished project to add a variable to data buttons

local vSeeds, vTrees, vBulbs, vRedFlag

if vSeeds="X"
    Field S
    S=0
    Field RedFlag
    vRedFlag=RedFlag
endif 
___ ENDPROCEDURE .NoCatalogsMacros _____________________________________________

___ PROCEDURE CandiansNoProvenceCode ___________________________________________



local vCanada, vRecCount
vRecCount=info("records")

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
if vRecCount=info("selected")
    message "No Canadians without Provence Codes"
endif
___ ENDPROCEDURE CandiansNoProvenceCode ________________________________________

___ PROCEDURE .ExportMacros ____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros from " +info("databasename")+" are saved to your clipboard!"
___ ENDPROCEDURE .ExportMacros _________________________________________________
