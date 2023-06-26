___ PROCEDURE bulbs_exporter/b _________________________________________________
global order, thiswindow, vord1, vord2, foldername, parentfoldername

foldername=""
thiswindow =   info("WindowName") 
openfile "45bulbs lookup"
openfile "bulbextractor"
deleteall
window thiswindow
field OrderNo
sortup
selectwithin OrderNo > 500000 AND OrderNo < 600000
selectwithin Order notcontains "0000" and Order ≠ ""
firstrecord

noshow

    Loop
        ;;ArrayFilter Order,order,¶,str(OrderNo)+¬+extract(Order,¶,seq())+¬+str(Sequence)
        ArrayFilter Order,order,¶,str(OrderNo)+¬+arrayrange(extract(Order,¶,seq()),1,8,¬)+¬+str(Sequence)
        window "bulbextractor:secret"
        openfile "+@order"
        window thiswindow
        downrecord
    until info("Stopped")
    
    endnoshow
    
openfile "bulbextractor"


firstrecord
if OrderNo=0
    deleterecord
endif
vord1=int(OrderNo)-500000
LastRecord
vord2=int(OrderNo)-500000

field Count
fill zeroblank(0)

save

parentfoldername=folderpath(dbinfo("folder",""))
foldername=folderpath(dbinfo("folder","")) +str(vord1)+"-"+str(vord2)

makefolder foldername


;;;;;;;;;;;;;;;;;;

Field Count
select Size = "A"
formulafill val(lookup("45bulbs lookup","number",Item,"size A","",0))
select Size = "B"
formulafill val(lookup("45bulbs lookup","number",Item,"size B","",0))
select Size = "C"
formulafill val(lookup("45bulbs lookup","number",Item,"size C","",0))
select Size = "D"
formulafill val(lookup("45bulbs lookup","number",Item,"size D","",0))
selectall


Field TotalCount
formulafill Count*Qty
Field RunningTotal
formulafill TotalCount

Field "Item"
SortUp
GroupUp
Field "Sequence"
SortUp

Field RunningTotal
runningtotal
removeallsummaries

Field "CountOrderedFromSuppliers"
FormulaFill lookup("bulbs purchased","Item",Item,"Qty",0,0)
Select RunningTotal <= CountOrderedFromSuppliers

Field "Item"
GroupUp

Field "Sequence"
Maximum

Field LastSeqToFill
formulafill lookup(info("databasename"),"Item",Item,"Sequence",0,1)

removeallsummaries
Selectall
Field "Item"
GroupUp

Field LastSeqToFill
maximum

Field CountOrderedFromSuppliers
maximum

Field RunningTotal
propagate

CollapseToLevel "1"
SelectWithin CountOrderedFromSuppliers > RunningTotal

Field "LastSeqToFill"
Fill 9999

SelectAll

SaveACopyAs foldername + ":45bulb running totals " +str(vord1)+"-"+str(vord2)

debug

Revert
removeallsummaries

;;;;;;;;;;;;;;;;;;;;;;;;
    SaveAs foldername + ":45bulb collation " +str(vord1)+"-"+str(vord2)
    
        field OrderNo
        GroupUp
        Field Total
        Total
        OutlineLevel "1"
        
                

        SaveACopyAs foldername + ":45bulb subtotals " +str(vord1)+"-"+str(vord2)
        Revert
        Field Size
        GroupUp
        Field Item
        GroupUp
        Field Qty
        Total
        Field Size
        Propagate
        RemoveDetail "Data"
        RemoveSummaries 1
        RemoveSummaries 2
        Field Item
        Sortup
              
        
        SaveAs foldername + ":45bulb packets "+str(vord1)+"-"+str(vord2)
            ;; look for itemtotals in the parent folder
            
            debug
                                    
            OpenFile parentfoldername + "45bulbs itemtotals"
            OpenFile "&" + foldername + ":45bulb packets " +str(vord1)+"-"+str(vord2)
            
            ;; make itemtotals save in the new folder
            Call itemtotals
 
___ ENDPROCEDURE bulbs_exporter/b ______________________________________________

___ PROCEDURE .Initialize ______________________________________________________
global n, raya, rayb, groupArray, rayc, rayd, raye, rayf, rayg, rayh, rayi, conArray, memArray, waswindow,
Com, Cost, Disc, VDisc, ODisc, stax,item, size,ono, vfax, state, sub, vzip,
vd, adj, tax, f, gr, di, da,  three, secno, alt, oono, nu, oldtot, newtot, addressArray, taxable,
mailaddress, mailcopies, mailheader, messageBody, newyear, newpfile, rollingdisc,ID, taxable
expressionstacksize 75000000
ID=0
ono=0
oono=0
nu=0
n=1
raya=¶
rayb=""
groupArray=""
rayc=""
rayd=""
raye=""
rayf=""
rayg=""
rayh=""
rayi=0
mailaddress=""
mailcopies=""
mailheader=""
messageBody=""
newyear="45"
conArray=""
memArray=""
;; newpfile=?(today()<date("12/1/22"),"44","45")

call .AutomaticFY

case folderpath(dbinfo("folder","")) contains "/ogs"
    Select (OrderNo >= 300000 and OrderNo < 400000)
case folderpath(dbinfo("folder","")) contains "/seeds"
    Select (OrderNo >= 700000 and OrderNo < 1000000)
case folderpath(dbinfo("folder","")) contains "/bulbs" or folderpath(dbinfo("folder","")) contains "FB"
    Select (OrderNo >= 500000 and OrderNo < 600000)
case folderpath(dbinfo("folder","")) contains "/trees"
    Select (OrderNo >= 400000 and OrderNo < 500000)
case folderpath(dbinfo("folder","")) contains "/mt"
    Select (OrderNo >= 600000 and OrderNo < 700000)    
endcase

field OrderNo 
sortup
___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE .addtoorder ______________________________________________________
arrayfilter rayh, rayh,¶, arrayinsert(extract(rayh,¶,seq()),1,1,¬)
;arrayfilter rayh, rayh,¶, arrayinsert(extract(rayh,¶,seq()),3,1,¬)
arrayfilter rayh, rayh,¶, arrayinsert(extract(rayh,¶,seq()),5,2,¬)
Order=Order+¶+rayh
Notes2="Added to order: "+rayh+¶+Notes2
rayh=""
call ".refigure"
___ ENDPROCEDURE .addtoorder ___________________________________________________

___ PROCEDURE .chargedeclined __________________________________________________
local badcard
badcard=""
waswindow=info("windowname")
Paid=0
«1stPayment»=0
«BalDue/Refund»=Paid-GrTotal
ShipCode="D"
field «Notes1»
if «Notes1»=""
«Notes1»="Card declined"
else field «Notes2»
if «Notes2»=""
«Notes2»="Card declined"
else «Notes3»="Card declined"
endif
endif
ArrayLineBuild badcard,¶,info("databasename"),¬+str(«C#»)+¬+str(«OrderNo»)+¬+Group+¬+Con+¬+MAd+¬+City+¬+
    St+¬+pattern(Zip,"#####")+¬+Telephone+¬+Email+¬+str(GrTotal)+¬+CreditCard+¬+ExDate
openfile "problem orders"
openfile "+@badcard"
Reason="Card Declined"
window waswindow




___ ENDPROCEDURE .chargedeclined _______________________________________________

___ PROCEDURE .checkno _________________________________________________________
if val(CreditCard[1,1])≥4 and val(CreditCard[1,1])≤6 and length(CreditCard)≠16 
field CreditCard
beep
editcell
endif
if val(CreditCard[1,1])=3 and length(CreditCard)≠15
field CreditCard
beep
editcell
endif
if val(CreditCard[1,1])<3 or val(CreditCard[1,1])>6
field CreditCard
beep
editcell
endif
call ".creditcheck"
field ExDate
___ ENDPROCEDURE .checkno ______________________________________________________

___ PROCEDURE .closewindow _____________________________________________________
CloseWindow
___ ENDPROCEDURE .closewindow __________________________________________________

___ PROCEDURE .creditcheck _____________________________________________________
local cctemp, ccvalid, invalidcard
waswindow=info("windowname")
if CreditCard =""
stop
endif
checkagain:
cctemp=CreditCard
cardvalidate cctemp, ccvalid
invalidcard=""
if ccvalid <>"Ok"
    YesNo "invalid card number, try again"
        if clipboard()="No"
        ArrayLineBuild invalidcard,¶,info("databasename"),¬+str(«C#»)+¬+str(«OrderNo»)+¬+Group+¬+Con+¬+MAd+¬+City+¬+
        St+¬+pattern(Zip,"#####")+¬+Telephone+¬+Email+¬+str(GrTotal)+¬+CreditCard+¬+ExDate
        window "problem orders"
        if error
        openfile "problem orders"
        endif
        openfile "+@invalidcard"
        Reason="Invalid card number"
        window waswindow
        CreditCard=""
        ShipCode="D"
        field «Notes1»
            if «Notes1»=""
            «Notes1»="Invalid Card"
            else field «Notes2»
                if «Notes2»=""
                «Notes2»="Invalid Card"
                else «Notes3»="Invalid Card"
                endif
            endif
        goto end
        endif
    field CreditCard
    editcell
    goto checkagain
endif
end:
field ExDate
___ ENDPROCEDURE .creditcheck __________________________________________________

___ PROCEDURE .creditex ________________________________________________________
if info("trigger")="Button.Charge"
waswindow=info("windowname")
Global credit
If CreditCard≠""
if AddPay1=GrTotal
call "next/1"
else
arraylinebuild credit, ¬, info("databasename"),¬+str(«OrderNo»)+¬+pattern(Zip,"#####")+¬+Email
    +¬+pattern(GrTotal,"#.##")+¬+¬+¬+CreditCard+¬+ExDate+¶
OpenFile "credit charges"
openfile "+@credit"
save
Windowtoback "credit charges"
window waswindow
Paid=GrTotal
AddPay1=Paid
DatePay1=today()
«BalDue/Refund»=0
endif
else 
message "They have no credit card"
Endif
endif
___ ENDPROCEDURE .creditex _____________________________________________________

___ PROCEDURE .cfill ___________________________________________________________
local vcc, vex
if ExDate=""
stop
else
endif
waswindow=info("windowname")
Num=«C#»
vcc=CreditCard
vex=ExDate
if «C#»=0
stop
endif
Window "customer_history:secret"
if «C#»≠Num
find «C#»=Num
endif
CreditCardNo=vcc
ExpDate=vex
window waswindow
vcc=""
vex=""

        

___ ENDPROCEDURE .cfill ________________________________________________________

___ PROCEDURE .coupon __________________________________________________________
;Notes3="ME2016"
;Discount=Discount+.1
;call .retotal
___ ENDPROCEDURE .coupon _______________________________________________________

___ PROCEDURE .custnumber ______________________________________________________
//this runs when C# is changed in Orders
//users are currently using two enters to get this to run during order entry
waswindow=info("windowname")
Global Num
Num=«C#»
ono=OrderNo
aray=""
conArray=""

call ".IsInternetOrder"


case paper_order=-1
    //the following is legacy code kept because I'm only fixing internet orders
    if MAd≠""
        addressArray=MAd+"."+pattern(Zip,"#####")
    endif

    if Con≠""
        conArray=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1]
    endif

    if «C#»=0
        window newyear+" mailing list"
        call ".newzip"
        stop
    endIf 

    window newyear+" mailing list"
    find «C#»=Num

    if info("found")=0
        call "getzip/Ω"
    else
        window waswindow
        call ".customerfill"   
    endif

case internet_order=-1
    call "NewSearch/ß"

endcase
    


___ ENDPROCEDURE .custnumber ___________________________________________________

___ PROCEDURE .customerfill ____________________________________________________
local Flag
Flag=""

waswindow=info("windowname") //XXorders
///this could need a call to info("formname") for the cases below

call ".FromBranch"

Num=«C#»

//Make sure you can find the new Number in the Mailing List
window thisFYear+" mailing list"
find «C#»=Num
if info("found")=0
    case paper_order=-1
        //old macro, keeping for paper, for now
            call "getzip/Ω"
            stop
    case internet_order=-1
            call "NewSearch/ß"
    endcase
endif

rayj=«Mem?»

«M?»=?(«M?» contains "E" or «M?» contains "U" or «M?» contains "R", "", «M?»)

//____Is it time to Update their RedFlag? e.g. It says "possible bad address"
if RedFlag≠""
    yesno "Current Redflag: "+RedFlag+". Erase this?"
    if clipboard()="Yes"
        RedFlag=""
    endif
endif

//____Assigns catalogs to receive based on order____//
if RedFlag=""
    case fromBranch contains "bulbs"
            Bf=?(Bf=0,1,Bf)
    case fromBranch contains "seeds" or fromBranch contains "ogs"
            S=?(S=0,1,S)
    case fromBranch contains "trees"
            T=?(T=0,1,T)
    endcase
endif

//******!!needs testing
//SpareText3=datepattern(today(),"MM/DD/YY")+" - Most Recent Order"+" "+"S"+str(S)+" "+"T"+str(T)+" "+"B"+str(Bf)
//currently is holding "Previous Addresses" from an incomplete macro
//perhaps make a dicitonary of both of these data points?
//___update__________
//dictionary made that holds last bulk mailing, could probably hold this as well


If Outstanding>0
    Message "This customer owes us from last year."
    stop
endif

window waswindow
window "customer_history:secret"
find «C#»=Num
window waswindow
Flag=grabdata(thisFYear+" mailing list", RedFlag)

If Flag≠""
    case Flag="changed"
        Flag=""
    case Flag="new"
        Flag=""
    case Flag contains "bad" or Flag contains "returned" or Flag contains "moved" or Flag contains "no forward"
        Message "Check this order carefully."
    endcase
    
    window thisFYear+" mailing list"
    RedFlag=Flag
    window waswindow
EndIf

;; this is just trying to check whether it's an internet order
if internet_order≠-1 or paper_order≠-1
    call ".IsInternetOrder"
endif

if internet_order=-1
    yesno "Update Con, Group, Addresses, Phone, and Email automatically from this order to the Mailing List?"
        if clipboard()="Yes"
            window thisFYear+"mailing list"
                «Group»=grabdata(thisFYear+"orders",«Group»)
                Con=grabdata(thisFYear+"orders", Con)
                MAd=grabdata(thisFYear+"orders", MAd)
                City=grabdata(thisFYear+"orders", City)
                St=grabdata(thisFYear+"orders", St)
                Zip=grabdata(thisFYear+"orders", Zip)
                SAd=grabdata(thisFYear+"orders", SAd)
                Cit=grabdata(thisFYear+"orders", Cit)
                Sta=grabdata(thisFYear+"orders", Sta)
                Z=grabdata(thisFYear+"orders",Z)
                phone=grabdata(thisFYear+"orders",Telephone)
                email=grabdata(thisFYear+"orders", Email)
        endif
    window waswindow
    goto GetLastYear
endif

UpdateFromMailingList:

    Group=grabdata(thisFYear+" mailing list",Group)
    Con=grabdata(thisFYear+" mailing list", Con)
    MAd=grabdata(thisFYear+" mailing list", MAd)
    City=grabdata(thisFYear+" mailing list", City)
    St=grabdata(thisFYear+" mailing list", St)
    Zip=grabdata(thisFYear+" mailing list", Zip)
    SAd=grabdata(thisFYear+" mailing list", SAd)
    Cit=grabdata(thisFYear+" mailing list", Cit)
    Sta=grabdata(thisFYear+" mailing list", Sta)
    Z=grabdata(thisFYear+" mailing list",Z)
    Telephone=grabdata(thisFYear+" mailing list",phone)
    Email=grabdata(thisFYear+" mailing list", email)
    ;comgrower=grabdata(thisFYear+" mailing list", CG)


GetLastYear:

    case info("formname")="seedsinput"
        LastYearTotal=grabdata("customer_history", ("S"+lastFYear))
    case info("formname")="bulbsinput"
        LastYearTotal=grabdata("customer_history", ("Bf"+lastFYear))
    ;case info("formname")="ogsinput"
        ;LastYearTotal=grabdata("customer_history", OGS43)
    case info("formname")="treesinput"
        LastYearTotal=grabdata("customer_history", ("T"+lastFYear))
    case info("formname")="mtinput"
        LastYearTotal=grabdata("customer_history", ("M"+lastFYear))
    endcase

«C#Text»=str(«C#»)

If Taxable="N"
    window thisFYear+" mailing list:secret"
    TaxEx="Y"
    window waswindow
endif

;; exclude scionwood order number range
If info("formname")="treesinput" And Z≠ 0 and (OrderNo < 410000 or OrderNo > 420000)
    case ShipCode ="T"
        Pool = 20
    case ShipCode ="P"
        Pool = 22
    defaultcase
        call ".pool"
    endcase
endif

if rayj="Y"
    call ".memberdisc"
else
    MemDisc = 0
    call ".retotal"
endif

Field SAd

If SAd=""
    stop
else
    Field Telephone
    If Telephone≠""
        field Email
        if Email≠""
            if info("formname")="ogsinput"
                call .rollingdiscount
                Field «1stPayment»
            else
                Field Sub1
                editcell
                Field Sub2
                editcell
                field «1stPayment»
            endIf
        endif
        EditCellStop
     else
        if info("formname")="ogsinput"
            Field «1stPayment»
        else
            field Sub1
            editcell
            field Sub2
        endIf
    endif
Endif 
___ ENDPROCEDURE .customerfill _________________________________________________

___ PROCEDURE .currentrecord ___________________________________________________
If OrderNo≠int(OrderNo)
    UpRecord
    groupArray=?(Group≠"", Group, Con)
    Downrecord
    Group=groupArray
    field Con
    stop
else
    stop
endif
___ ENDPROCEDURE .currentrecord ________________________________________________

___ PROCEDURE .depotshipping ___________________________________________________
global depname, deplist
depname=""
deplist=""

if ShipCode <> "J"
    stop
endif

waswindow=info("windowname")
openfile "shipdepotlookup"
window waswindow
arraybuild deplist, ¶, "shipdepotlookup", «depot_order_num»+¬+depot_name
deplist=strip(deplist)
deplist=arraystrip(deplist,¶)
popup deplist, 144, 72, "MELI21", depname
depname=extract(depname,¬,1)
window "shipdepotlookup:secret"
find depot_order_num=depname
ono=price_per_pound
message ono
window waswindow
OrderComments=depname+¶+OrderComments
«$Shipping»=ShippingWt*ono
«$Shipping»=?(«$Shipping»<3.00,3.00,«$Shipping»)

call .retotal
___ ENDPROCEDURE .depotshipping ________________________________________________

___ PROCEDURE .groupemail ______________________________________________________
local orderarray, piecearray
piecearray=""
orderarray=""
ono=OrderNo
field OrderNo
sortup
find OrderNo=ono
downrecord
loop
    piecearray=replace(Order, ¬, " ")
    ;piecearray=replace(Order, ¶, "; ")
    piecearray=piecearray+¶+"Subtotal: "+str(Subtotal)+¶
    orderarray=orderarray+str(OrderNo)[".",-1][2,-1]+¶+piecearray
    downrecord
    stoploopif OrderNo≥ono+1
until info("stopped")

find OrderNo=ono
orderarray=orderarray+¶+"Order Subtotal= "+str(Subtotal)+¶+?(SalesTax>0,"Sales Tax=" +str(SalesTax)+¶+"Adjusted Total= "+str(AdjTotal)+
    ¶,"")+?(VolDisc>0, "Volume Discount= "+str(VolDisc)+¶,"")+?(«$Shipping»>0, "Shipping= "+str(«$Shipping»)+¶,"")+"Total= "+str(OrderTotal)

superobject "emailbody", "open"
activesuperobject "setselection", 0, 32768
activesuperobject "clear"
activesuperobject "inserttext", orderarray
activesuperobject "close"

if Email≠''
    mailaddress=Email
    mailheader= "Your Moose Tuber Order "+str(OrderNo)
    superobject "emailaddress", "open", "inserttext", mailaddress
    activesuperobject "close"
endif

superobject "emailsubject", "open", "inserttext", mailheader
activesuperobject "close"
___ ENDPROCEDURE .groupemail ___________________________________________________

___ PROCEDURE .input ___________________________________________________________
If info("trigger")="Button.input"
    closewindow
    case OrderNo > 300000 and OrderNo < 400000
        goform "ogsinput"
    case OrderNo > 400000 and OrderNo < 500000
        goform "treesinput"
    case OrderNo > 500000 and OrderNo < 600000
        goform "bulbsinput"
    case OrderNo > 600000 and OrderNo < 700000
        goform "mtinput"
    case OrderNo > 700000 and OrderNo < 1000000
        goform "seedsinput"
    endcase
endif
___ ENDPROCEDURE .input ________________________________________________________

___ PROCEDURE .InvalidCard _____________________________________________________
        waswindow=info("windowname")
        ArrayLineBuild invalidcard,¶,info("databasename"),¬+str(«C#»)+¬+str(«OrderNo»)+¬+Group+¬+Con+¬+MAd+¬+City+¬+
        St+¬+pattern(Zip,"#####")+¬+Telephone+¬+Email+¬+str(GrTotal)+¬+CreditCard+¬+ExDate
        window "problem orders"
        
        if error
            openfile "problem orders"
        endif
        
        openfile "+@invalidcard"
        Reason="Invalid card number"
        window waswindow
        ShipCode="D"
        Paid=0
        «1stPayment»=0
        «BalDue/Refund»=Paid-GrTotal
___ ENDPROCEDURE .InvalidCard __________________________________________________

___ PROCEDURE .memberdisc ______________________________________________________
case OrderNo > 700000 and OrderNo<1000000
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=float(Subtotal*.01)
    AdjTotal=Subtotal-VolDisc-MemDisc

    If Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif

    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»+Donation+Membership
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    
case OrderNo > 300000 and OrderNo < 400000
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=float(Subtotal*.01)
    AdjTotal=Subtotal-VolDisc-MemDisc

    If Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif

    OrderTotal=AdjTotal+SalesTax+«$Shipping»
    GrTotal=«OrderTotal»+Donation
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    
case OrderNo > 600000 and OrderNo < 700000
    VolDisc=float(«4SpareMoney»)*float(Discount)
    MemDisc=float(Subtotal*.01)
    AdjTotal=Subtotal-VolDisc-MemDisc

    If Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif

    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»+Donation
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    
case OrderNo > 400000 and OrderNo < 500000
    memArray=?(«$Shipping»<18, "Y", "N")
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=float(Subtotal*.01)
    AdjTotal=Subtotal-VolDisc-MemDisc
    
    If Taxable="Y"
        // enabled salestax recalculation for refiguring orders.
        call ".salestax"
    endif
    
    if «$Shipping»>10.00 and MemDisc>0
        call ".refiguretreeshipping"
    endif
    
    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»+Donation+«Donation2»
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    
case OrderNo > 500000 and OrderNo < 600000
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=float(Subtotal*.01)
    AdjTotal=Subtotal-VolDisc-MemDisc
    BulbsAdjTotal=BulbsSubtotal-(VolDisc+MemDisc)*divzero(BulbsSubtotal,Subtotal)

    If Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif
 
    if ShipCode contains "T" or ShipCode contains "P"
        «$Shipping»=0
    else
    ;; TODO find out what actual shipping rates are
        «$Shipping»=?(BulbsAdjTotal>100, BulbsAdjTotal*.12,12)
    
        If Sta contains "HI" or Sta contains "AK" or Sta contains "APO"
            «$Shipping»=?(BulbsAdjTotal>100, BulbsAdjTotal*.16,16)
        endif
    endif
    
    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
endcase
___ ENDPROCEDURE .memberdisc ___________________________________________________

___ PROCEDURE .newzip __________________________________________________________
global listzip, thiszip, findher, findzip, findcity, newcity, findname,findname1, findname2, thisname
serverlookup "off" 
waswindow=info("windowname")
listzip=""
thiszip=""
newcity=""
again:
findher=""
supergettext findher, {caption="Enter Address,Zip" height=100 width=400 captionfont=Times captionsize=14 captioncolor="cornflowerblue"
    buttons="Find;Redo;Cancel"}
    if info("dialogtrigger") contains "Find"
    findzip=extract(findher,",",2)
    findzip=strip(findzip)
    findcity=extract(findher,",",1)
    liveclairvoyance findzip,listzip,¶,"",newyear+" mailing list",pattern(Zip,"#####"),"=",Con+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")+¬+phone,0,0,""
    arraysubset listzip, listzip, ¶, import() contains findcity
    if info("found")=0
    goto tryname
    endif
    endif
    if info("dialogtrigger") contains "Redo"
    findher=""
    goto again
    endif
    if info("dialogtrigger") contains "Cancel"
    stop
    endif

;selectall
window newyear+" mailing list"
superchoicedialog listzip, thiszip, {height=400 width=500 font=Helvetica caption="Click on one and then hit OK or New for new entry" 
        captionfont=Times captionsize=12 captioncolor=red size=14 buttons="OK:100;New:100;Cancel:100"}
        
if info("dialogtrigger") contains "OK"
    message thiszip
    find MAd=extract(thiszip, ¬,2) and City contains extract(thiszip, ¬,3)
    stop
    call "enter/e"
endif

if info("dialogtrigger") contains "New"
    gettext "Which town?", newcity
    if newcity≠""
        ;window newyear+" mailing list"
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

tryname:
 window waswindow
    findname=""
    findname1=""
    findname2=""
    supergettext findname, {caption="Enter First and Last Name" height=100 width=400 captionfont=Times captionsize=14 captioncolor="limegreen"
    buttons="Find;Redo;Cancel"}
    if info("dialogtrigger") contains "Find"
        findname1=findname[1," "]
        findname1=strip(findname1)
        findname2=findname[" ",-1]
        findname2=strip(findname2)
        liveclairvoyance findname,thisname,¶,"",newyear+" mailing list",pattern(Zip,"#####"),"=",Con+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")+¬+phone,0,0,""
    endif
    
    if info("dialogtrigger") contains "Redo"
        goto tryname
    endif
    
    if info("dialogtrigger") contains "Cancel"
        stop
    endif
    
    superchoicedialog findname, thisname, {height=400 width=500 font=Helvetica caption="Click on one and then hit OK or New for new entry" 
        captionfont=Times captionsize=12 captioncolor=green size=14 buttons="OK:100;New:100;Cancel:100"}
        
     if info("dialogtrigger") contains "OK"
        window newyear+" mailing list"
        find Con=extract(thisname, ¬,1) and City contains extract(thisname, ¬,3)
        call "enter/e"
     endif  
     
    if info("dialogtrigger") contains "New"
        gettext "Which town?", newcity
        if newcity≠""
            window newyear+" mailing list"
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

___ ENDPROCEDURE .newzip _______________________________________________________

___ PROCEDURE .opengiftcertificate _____________________________________________
shellopendocument "http://www.fedcoseeds.com/manage_site/gift-certificates"
___ ENDPROCEDURE .opengiftcertificate __________________________________________

___ PROCEDURE .openOrder _______________________________________________________
if info("formname") contains "input"
setwindowrectangle rectanglesize(72,33,876,805),""
openform "order"
endif
___ ENDPROCEDURE .openOrder ____________________________________________________

___ PROCEDURE .salestax ________________________________________________________
/* New in 44: 4 arrays for sales tax states:

taxstates = all items taxable, shipping taxed
noship = all items taxable, shipping not taxed
tshipexemptions = some items taxable, that portion of shipping taxed
noshipexemptions = some items taxable, shipping not taxed

These states are correct for Trees only. Other branches may have differences in exemptions.
Old macro below!
*/

local taxstates, tshipexemptions, noship, noshipexemptions, taxitems, taxtotal
;; TREES
if OrderNo>400000 and OrderNo<500000
    taxstates="CT,GA,IL,IN,KS,KY,MI,MN,NC,NJ,NM,NY,OH,PA,RI,TN,WA,WI,WV"
    tshipexemptions="VT"
    noship="FL,MD,ME,UT,VA,WY"
    noshipexemptions="MA"

    taxable=TaxTotal*(1-Discount-?(MemDisc > 0,.01,0))

    case arraycontains(taxstates+","+tshipexemptions+","+noship+","+noshipexemptions,TaxState,",")=0
        TaxedAmount=0

    case arraycontains(noship,TaxState,",")=-1
        TaxedAmount=taxable

    case arraycontains(noshipexemptions,TaxState,",")=-1
        taxable=(TaxTotal-?(VolDisc≠0,round(float(TaxTotal)*float(Discount)+.0001,.01),0)
            -?(MemDisc≠0,round((TaxTotal*0.01)+.0001,.01),0))
        TaxedAmount=taxable

    case arraycontains(taxstates,TaxState,",")=-1
        TaxedAmount=taxable+Surcharge+«$Shipping»

    case arraycontains(tshipexemptions,TaxState,",")=-1
        taxable=(TaxTotal-?(VolDisc≠0,round(float(TaxTotal)*float(Discount)+.0001,.01),0)
            -?(MemDisc≠0,round((TaxTotal*0.01)+.0001,.01),0)+round(«$Shipping»*float(TaxTotal/Subtotal)+.0001,.01))
        TaxedAmount=taxable

    endcase

    TaxedAmount=?(Taxable="Y",TaxedAmount,0)
endif

;; all other branches use a different definition of taxstates and noship than Trees does, 
;; so these variables get reassigned outside of the conditional.

taxstates="CT,GA,IL,IN,KY,KS,MI,MN,NC,NJ,NM,NY,OH,PA,RI,TN,VT,WA,WI,WV"
noship="FL,MA,MD,ME,VA,UT,WY"

;; SEEDS
if OrderNo > 700000 and OrderNo < 1000000
    case TaxState contains "VT" or TaxState contains "RI" or TaxState contains "CT"
        field Order
        arrayfilter Order, taxitems,¬,?(val(extract(extract(Order,¶,seq()),¬,2))>4700,extract(extract(Order,¶,seq()),¬,8),"")
        taxitems=arraystrip(taxitems,¬)
        taxtotal=arraynumerictotal(taxitems,¬)
        TaxedAmount=?(Taxable="Y", taxtotal*float(divzero(AdjTotal,Subtotal))+float(«$Shipping»*float(divzero(taxtotal,Subtotal))),0)
        if TaxState contains "MD"
            TaxedAmount=?(Taxable="Y",taxtotal*float(divzero(AdjTotal,Subtotal)),0)
        endif
        if TaxRate = 0
            TaxRate = lookup("ZipCodeList","ZipCode",Z,"TaxRate",0,0)
            StateRate = lookup("ZipCodeList","ZipCode",Z,"StateRate",0,0)
            CountyRate = lookup("ZipCodeList","ZipCode",Z,"CountyRate",0,0)
            CityRate = lookup("ZipCodeList","ZipCode",Z,"CityRate",0,0)
            SpecialRate = lookup("ZipCodeList","ZipCode",Z,"SpecialRate",0,0)
        endif

    case arraycontains(taxstates+","+noship,TaxState,",")=0
        TaxedAmount=0

    case arraycontains(taxstates,TaxState,",")=-1
        TaxedAmount=?(Taxable="Y",AdjTotal+Surcharge+«$Shipping»,0)

    case arraycontains(noship,TaxState,",")=-1
        if TaxState contains "MA" or TaxState contains "MD"
            field Order
            arrayfilter Order, taxitems,¬,?(val(extract(extract(Order,¶,seq()),¬,2))>4700,extract(extract(Order,¶,seq()),¬,8),"")
            taxitems=arraystrip(taxitems,¬)
            taxtotal=arraynumerictotal(taxitems,¬)
            TaxedAmount=?(Taxable="Y", taxtotal*float(divzero(AdjTotal,Subtotal)),0)
            if TaxRate = 0
                TaxRate = lookup("ZipCodeList","ZipCode",Z,"TaxRate",0,0)
                StateRate = lookup("ZipCodeList","ZipCode",Z,"StateRate",0,0)
                CountyRate = lookup("ZipCodeList","ZipCode",Z,"CountyRate",0,0)
                CityRate = lookup("ZipCodeList","ZipCode",Z,"CityRate",0,0)
                SpecialRate = lookup("ZipCodeList","ZipCode",Z,"SpecialRate",0,0)
            endif
        else
            TaxedAmount=?(Taxable="Y",AdjTotal,0)
        endif
    endcase
endif



;; OGS
if OrderNo > 300000 and OrderNo < 400000
    taxable=TaxTotal*(1-Discount-?(MemDisc > 0,.01,0))

    case arraycontains(taxstates+","+noship,TaxState,",")=0
        TaxedAmount=0
    case arraycontains(noship,TaxState,",")=-1
        TaxedAmount=taxable
    case arraycontains(taxstates+","+noship,TaxState,",")=-1
        TaxedAmount=taxable+divzero(taxable,AdjTotal)*«$Shipping»
    endcase

    TaxedAmount=?(Taxable="Y",TaxedAmount,0)
endif

;; BULBS
if OrderNo > 500000 and OrderNo < 600000
    taxable=taxable*(1-Discount-?(MemDisc > 0,.01,0))

    case arraycontains(taxstates+","+noship,TaxState,",")=0
        ;; states in which we do NOT have nexus
        TaxedAmount=0
    case arraycontains(noship,TaxState,",")=-1
        TaxedAmount=taxable
    case arraycontains(taxstates+","+noship,TaxState,",")=-1
        TaxedAmount=taxable+divzero(bulbstaxtotal,BulbsSubtotal)*«$Shipping»
    endcase

    TaxedAmount=?(Taxable="Y",TaxedAmount,0)
endif

;; POE -- UNTESTED, NEEDS TO BE TESTED FOR FY 45 WHEN THERE ARE ACTUAL ORDERS
if OrderNo > 600000 and OrderNo < 700000
    taxable=TaxTotal*(1-Discount-?(MemDisc > 0,.01,0))

    case arraycontains(taxstates+","+noship,TaxState,",")=0
        TaxedAmount=0
    case arraycontains(noship,TaxState,",")=-1
        TaxedAmount=taxable
    case arraycontains(taxstates+","+noship,TaxState,",")=-1
        TaxedAmount=taxable+divzero(taxable,AdjTotal)*«$Shipping»
    endcase

    if arraycontains("CT MA MD MN NY RI VT", TaxState," ")= -1
        if TaxRate = 0
            TaxRate = lookup("ZipCodeList","ZipCode",Z,"TaxRate",0,0)
            StateRate = lookup("ZipCodeList","ZipCode",Z,"StateRate",0,0)
            CountyRate = lookup("ZipCodeList","ZipCode",Z,"CountyRate",0,0)
            CityRate = lookup("ZipCodeList","ZipCode",Z,"CityRate",0,0)
            SpecialRate = lookup("ZipCodeList","ZipCode",Z,"SpecialRate",0,0)
        endif
    endif

    TaxedAmount=?(Taxable="Y",TaxedAmount,0)
endif

SalesTax=float(TaxedAmount)*float(TaxRate)
StateTax=float(TaxedAmount)*float(StateRate)
CountyTax=float(TaxedAmount)*float(CountyRate)
CityTax=float(TaxedAmount)*float(CityRate)
SpecialTax=float(TaxedAmount)*float(SpecialRate)

___ ENDPROCEDURE .salestax _____________________________________________________

___ PROCEDURE .Group ___________________________________________________________
waswindow=info("windowname")

if info("formname") contains "addresschecker"
    stop
endif

if info("trigger")="Key.Tab"
    Field «Con»
endIf

ID=«C#»

if info("trigger")="Key.Return"
    ;message newyear
    window newyear+" mailing list:secret"
    
    if «C#»≠ID
    
        if info("selected")<info("records")
            selectall
        endif
        
        find «C#»=ID
    endif
    
    Group=grabdata(newyear+"orders", Group)
    window "customer_history:secret"
    Find «C#»=ID
    Group=grabdata(newyear+"orders", Group)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    field «Con»
EndIf


___ ENDPROCEDURE .Group ________________________________________________________

___ PROCEDURE .Con _____________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
stop
endif
if info("formname") contains "addresschecker"
stop
endif
if OrderNo≠int(OrderNo)
       if info("windowname")≠"addresschecker"
            ;call ".dropship"
            if info("formname")="treesinput"
            Pool=rayi
            endif
            if info("windowname")≠"ogsinput"
            field Sub1
            editcell
            field Sub2
            editcell
            endif
            Downrecord
                if OrderNo≠int(OrderNo)
                    if info("formname")="treesinput"
                        Pool=rayi
                    endif
                    if info("formname")≠"addresschecker"
                        Group=groupArray
                    else Group=Group
                    endif
                    MemDisc=?(rayg="Y", .01, MemDisc)
                    field Con
                    editcellstop
                else
                    field «C#»
                    editcellstop
                endif
    endif              
    if info("windowname")≠"ogsinput"
        Sub1="Y"
      
        downrecord
        field Con
    endif
    if info("windowname")="addresschecker"
        UpRecord
        groupArray=?(Group≠"", Group, Con)
        Downrecord
        Group=groupArray
        stop
    endif
    if OrderNo=int(OrderNo)
        field «C#»
        if info("formname")="addresschecker"
            stop
        endif
        editcellstop
    endif
    stop
endif

If info("trigger")="Key.Tab"
    Field MAd
EndIf

ID=«C#»

If info("trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    Con=grabdata(newyear+"orders",Con)
    window "customer_history:secret"
    Find «C#»=ID
    Con=grabdata(newyear+"orders", Con)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field MAd
EndIf
___ ENDPROCEDURE .Con __________________________________________________________

___ PROCEDURE .MAd _____________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    Field City
EndIf
ID=«C#»
If info("Trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    MAd=grabdata(newyear+"orders", MAd)
    window "customer_history:secret"
    Find «C#»=ID
    MAd=grabdata(newyear+"orders", MAd)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field City
EndIf

___ ENDPROCEDURE .MAd __________________________________________________________

___ PROCEDURE .City ____________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    Field St
EndIf
ID=«C#»
If info("Trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    City=grabdata(newyear+"orders", City)
    window "customer_history:secret"
    Find «C#»=ID
    City=grabdata(newyear+"orders", City)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field City
    Field St
EndIf

___ ENDPROCEDURE .City _________________________________________________________

___ PROCEDURE .St ______________________________________________________________
debug

waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    Field Zip
EndIf
ID=«C#»
If info("Trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    St=grabdata(newyear+"orders", St)
    window "customer_history:secret"
    Find «C#»=ID
    St=grabdata(newyear+"orders", St)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field Zip
EndIf

___ ENDPROCEDURE .St ___________________________________________________________

___ PROCEDURE .Zip _____________________________________________________________
waswindow=info("windowname")
//don't do any of the following on addresschecker
    if info("formname") contains "addresschecker"
    stop
    endif
//don't do anything but move fields if I used Tab to get here
    If info("trigger")="Key.Tab"
    Field SAd
    EndIf

//CheckNum
if «C#»=0
    message "This Customer Must have a Customer # to use this address updater"
    stop
endif

//GetNumber
ID=«C#»

//________Update Files with New Address from Order______//
If info("trigger")="Key.Return"
YesNo "Update "+thisFYear+" "+"mailing list & customer_history Mailing Address?"
        If clipboard()="Yes"
      window thisFYear+" mailing list"
            If «C#»≠ID
            if info("selected")<info("records")
            selectall
            endif
            Find «C#»=ID
            EndIf
        MAd=grabdata(thisFYear+"orders", MAd)
        City=grabdata(thisFYear+"orders", City)
        St=grabdata(thisFYear+"orders", St)
        Zip=grabdata(thisFYear+"orders", Zip)
        adc=lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        Field «C#»
        window "customer_history:secret"
        Find «C#»=ID
        MAd=grabdata(thisFYear+"orders", MAd)
        City=grabdata(thisFYear+"orders", City)
        St=grabdata(thisFYear+"orders", St)
        Zip=grabdata(thisFYear+"orders", Zip)
        window waswindow
        EndIf
        if info("formname")=thisFYear+"orders:addresschecker"
stop
endif
Field SAd
EndIf

___ ENDPROCEDURE .Zip __________________________________________________________

___ PROCEDURE .SAd _____________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    Field Cit
EndIf
ID=«C#»
If info("trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    SAd=grabdata(newyear+"orders", SAd)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field Cit
EndIf

___ ENDPROCEDURE .SAd __________________________________________________________

___ PROCEDURE .Cit _____________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    Field Sta
EndIf
ID=«C#»
If info("Trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    Cit=grabdata(newyear+"orders", Cit)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field Sta
EndIf

___ ENDPROCEDURE .Cit __________________________________________________________

___ PROCEDURE .Sta _____________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    Field Z
EndIf
ID=«C#»
If info("Trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    Sta=grabdata(newyear+"orders", Sta)
    window waswindow
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field Z
EndIf

___ ENDPROCEDURE .Sta __________________________________________________________

___ PROCEDURE .Z _______________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
ID=«C#»
if ShipCode contains "P" or ShipCode contains "T"
    TaxState="ME"
    TaxRate=.055
else
    TaxState=Sta
    TaxRate=lookup("ZipCodeList", "ZipCode",«Z»,"TaxRate",0,0)
endif
YesNo "Update 45 mailing list with permanent SAd?"
If clipboard()="Yes"
    window newyear+" mailing list"
    If «C#»≠ID
        Find «C#»=ID
    EndIf
    If «UPS?»="√"
        NoYes "Are you sure you want to change?"
        If clipboard()="No"
            window waswindow
            Stop
        EndIf
    EndIf
    SAd=grabdata(newyear+"orders",SAd)
    Cit=grabdata(newyear+"orders",Cit)
    Sta=grabdata(newyear+"orders",Sta)
    Z=grabdata(newyear+"orders",Z)
    window waswindow
    if OrderNo=oono
        stop 
    endif
    Field Telephone
    If «1stPayment»>0 or «BalDue/Refund»<0
        stop
    endif
    If Telephone≠""
        case info("formname")="seedsinput"
            ;call ".seedpool"
            Field Sub1
            editcell
            field Sub2
            editcell
        case info("formname")="treesinput"
            call ".pool"
            Field Sub1
            editcell
            field Sub2
            editcell
        case info("formname")="bulbsinput"
            Field Sub1
            editcell
            field Sub2
            editcell
        case info("formname")="mtinput"
            Field Sub1
            editcell
            field Sub2
            editcell
        endcase
        field «1stPayment»
        stop
    else
        editcellstop
    EndIf
EndIf

If clipboard()= "No"
    Field Telephone
    If Telephone≠""
        case info("formname")="seedsinput"
            call ".seedpool"
        case info("formname")="treesinput"
            call ".pool"
        case info("formname")="bulbsinput"
        case info("formname")="mtinput"
        endcase
        if info("formname")=newyear+"orders:addresschecker"
            stop
        endif
        Field Sub1
        editcell
        field Sub2
        editcell
        field «1stPayment»
    else
        editcell
        stop
    EndIf
Endif
___ ENDPROCEDURE .Z ____________________________________________________________

___ PROCEDURE .phone ___________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    if OrderNo=oono
        stop 
    endif
    if info("formname")="ogsinput"
        field Email
        editcellstop
        field «1stPayment»
        ;editcell
        stop
    endif
    if info("formname")="treesinput"
        call ".pool"
    endif
    ; if info("formname")="seedsinput"
    ;call ".seedpool"
    ;endif
    field Email
    editcellstop
    Field Sub1
    editcell
    field Sub2
    editcell
    Field «1stPayment»
    stop
EndIf

ID=«C#»
If info("Trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    phone=grabdata(newyear+"orders",Telephone)
    window waswindow
    if OrderNo=oono
        stop 
    endif
    if info("formname")="ogsinput"
        field Email
        editcellstop
        field «1stPayment»
        ;editcell
        stop
    endif
    if info("formname")="treesinput"
        call ".pool"
    endif
    if info("formname")="seedsinput"
        call ".seedpool"
    endif
;if info("formname") contains "bulbs"
;Field «1stPayment»
;editcell
;stop
;Endif
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    field Email
    editcellstop
    Field Sub1
    editcell
    field Sub2
    editcell
    Field «1stPayment»
    stop
EndIf
___ ENDPROCEDURE .phone ________________________________________________________

___ PROCEDURE .email ___________________________________________________________
waswindow=info("windowname")
if info("formname") contains "addresschecker"
    stop
endif
If info("trigger")="Key.Tab"
    if OrderNo=oono
        stop 
    endif
    if info("formname")="ogsinput"
        field «1stPayment»
        ;editcell
        stop
    endif
    Field Sub1
    editcell
    field Sub2
    editcell
    Field «1stPayment»
    stop
EndIf
ID=«C#»
If info("Trigger")="Key.Return"
    window newyear+" mailing list:secret"
    If «C#»≠ID
        if info("selected")<info("records")
            selectall
        endif
        Find «C#»=ID
    EndIf
    email=grabdata(newyear+"orders",Email)
    window "customer_history:secret"
    Find «C#»=ID
    Email=grabdata(newyear+"orders", Email)
    window waswindow
    if OrderNo=oono
        stop 
    endif
    if info("formname")="ogsinput"
        field «1stPayment»
        ;editcell
        ;field TaxCharged
        ;editcell
        stop
    endif
    if info("formname")=newyear+"orders:addresschecker"
        stop
    endif
    Field Sub1
    editcell
    field Sub2
    editcell
    Field «1stPayment»
    ;editcell
    ;field TaxCharged
    stop
EndIf



___ ENDPROCEDURE .email ________________________________________________________

___ PROCEDURE .sub2 ____________________________________________________________
if info("formname") contains "addresschecker"
    stop
endif
field «1stPayment»
;editcell
stop
___ ENDPROCEDURE .sub2 _________________________________________________________

___ PROCEDURE .CG ______________________________________________________________
waswindow=info("windowname")
;if comgrower="Y"
;di="Y"
;else di="N"
;endif
window newyear+" mailing list:secret"
;CG=di
window waswindow
di=""
___ ENDPROCEDURE .CG ___________________________________________________________

___ PROCEDURE .pool ____________________________________________________________
local p1,p2,p3, p4,p5,p6,p7,p8,p9,p14,p50
p1="AL,AR,AZ,CA,FL,GA,HI,LA,MS,NC,SC,TN,TX"
p2="DC,DE,MD,MO,OK,VA,KY"
p3="IL,IN,NJ,NM,WV"
p4="OH,IA"
p5="KS,ND,NE,NV,OR,SD,WA"
p6="PA"
p7="CT,RI"
p8="MI"
p9="CO,ID,MT,UT,WY"
p14="WI,MN"
p50="AK"

if ShipCode="T"
Pool=90
else
if ShipCode="P"
    Pool=90

else
if «Notes2» contains "group order" or «1SpareText» contains "G"
if Sta≠"VT" and Sta≠"NH" and Sta≠"ME"
Pool=40
endif
if Sta="VT" or Sta="NH" or Sta="ME"
Pool=41
endif

else

// Southern NY
    Case Z>10000 and Z<12800
        Pool=10     
        
//Eastern MA
    Case Z>02000 and Z<02800
        Pool=11

//Central MA        
    Case Z>01400 and Z<02000
        Pool=12
        
//Western MA
    Case Z>01000 and Z<01400
        Pool=13

//Central/Southern VT
    Case Z>04999 and Z<05400
        Pool=15
        
//Northern NY
    Case Z>12799 and Z<15000
        Pool=16
        
//Southern NH
    Case Z>02999 and Z<03500
        Pool=17
        
        
//Central/Southern NH, Southern ME    
    Case (Z>03599 and Z<03898)
    or (Z>03899 and Z<04020)
        Pool=18
        
//Southern ME
    Case Z>04019 and Z<04100
        Pool=19
        
//Southern/Western ME
    Case Z>04099 and Z<04300
        Pool=20
        
//Midcoast ME
    Case (Z>04299 and Z<04400)
    or (Z>04499 and Z<04600)
        Pool=21
        
//Downeast ME
    Case Z>04599 and Z<04700
        Pool=22
        
//Central/Midcoast ME
    Case Z>04799 and Z≤04930
        Pool=23
        
//North/Central VT
    Case Z>05399 and Z<06000
        Pool=24
     
        
//Central/Western ME
    Case Z>04929 and Z<05000
        Pool=25
        
//Northern ME
    Case Z>04399 and Z<04500
        Pool=26
        
//Far Northern ME, Northern NH
    Case (Z>04699 and Z<04800)
    or (Z>03499 and Z<03600)
       Pool=27
       
       
Case arraycontains(p1,Sta,",")=-1
Pool=1
Case arraycontains(p2,Sta,",")=-1
Pool=2
Case arraycontains(p3,Sta,",")=-1
Pool=3
Case arraycontains(p4,Sta,",")=-1
Pool=4
Case arraycontains(p5,Sta,",")=-1
Pool=5
Case arraycontains(p6,Sta,",")=-1
Pool=6
Case arraycontains(p7,Sta,",")=-1
Pool=7
Case arraycontains(p8,Sta,",")=-1
Pool=8
Case arraycontains(p9,Sta,",")=-1
Pool=9
Case arraycontains(p14,Sta,",")=-1
Pool=14
Case arraycontains(p28,Sta,",")=-1
Pool=28
case ShipCode=D
Pool=0

        
    DefaultCase
        GetScrap "This zip code has no pool. Please assign."
        Pool=val(clipboard())
    EndCase
endif
endif
endif
___ ENDPROCEDURE .pool _________________________________________________________

___ PROCEDURE .rollingdiscount _________________________________________________
rollingdisc=""
global finddisccust
waswindow=info("windowname")
if info("files") notcontains "discounttable"
    openfile "discounttable"
    makesecret
endif
finddisccust=«C#»
window "discounttable:secret"
find «C#»=finddisccust
if info("found")=0
    rollingdisc="Not in discount table yet"
else
    if Bulk=1
        rollingdisc="Bulk Prices"
    else
        rollingdisc="Rolling Discount= "+str(Discount*100)+"%"
    endif
endif
window waswindow
«2SpareText»=rollingdisc
drawobjects
___ ENDPROCEDURE .rollingdiscount ______________________________________________

___ PROCEDURE .seedpool ________________________________________________________
Case val(pattern(Z,"#####")[1,3])≥090 and val(pattern(Z,"#####")[1,3])≤098
    Pool=1
Case val(pattern(Z,"#####")[1,3])≥962 and val(pattern(Z,"#####")[1,3])≤969
    Pool=1
Case val(pattern(Z,"#####")[1,3])≥995 and val(pattern(Z,"#####")[1,3])≤999
    Pool=1
Case val(pattern(Z,"#####")[1,3])=340
    Pool=1
DefaultCase
    Pool=0
EndCase
___ ENDPROCEDURE .seedpool _____________________________________________________

___ PROCEDURE .pottedmessage ___________________________________________________
local emailmessage
emailmessage=""
emailmessage=lookup("email messages","messagetitle",cannedmessage,"message","",0)
emailmessage=replace(emailmessage, "«Con»", Con)
emailmessage=replace(emailmessage, "«OrderNo»", str(OrderNo))
emailmessage=replace(emailmessage, "«GrTotal»", pattern(GrTotal, "$#.##"))
emailmessage=replace(emailmessage, "«$Shipping»", pattern(«$Shipping», "$#.##"))
superobject "emailbody", "open"
activesuperobject "setselection", 0, 32768
activesuperobject "clear"
activesuperobject "inserttext", emailmessage
activesuperobject "close"
___ ENDPROCEDURE .pottedmessage ________________________________________________

___ PROCEDURE .sendemail _______________________________________________________
local mailarray
mailarray=""
mailaddress=Email
mailheader= "Your Fedco Order "+str(OrderNo)
case mailcopies≠""
mailarray=mailaddress+","+mailcopies
sendarrayemail "", mailarray, mailheader, messageBody
case mailcopies=""
sendoneemail "", mailaddress, mailheader, messageBody
endcase
mailaddress=""
mailcopies=""
mailheader=""
messageBody=""
___ ENDPROCEDURE .sendemail ____________________________________________________

___ PROCEDURE .Refund __________________________________________________________
If info("Trigger") ="Button.All"
    If Paid=0
        Donated_Refund=GrTotal
    else
        Donated_Refund=Paid
    EndIf
Endif
«1stRefund»=Donated_Refund
___ ENDPROCEDURE .Refund _______________________________________________________

___ PROCEDURE .ogsbulk _________________________________________________________
global checkorder
waswindow=info("windowname")
checkorder=""
checkorder=Order
Bulk="bulk"
openfile "OrderRetotaller"
openfile "&@checkorder"
call "ogsbulkretotaller"
___ ENDPROCEDURE .ogsbulk ______________________________________________________

___ PROCEDURE .refigure ________________________________________________________
global checkorder, taxstate
 
waswindow=info("windowname")
checkorder=""
checkorder=Order
case OrderNo > 700000 and OrderNo < 1000000
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "seedsretotaller"
case OrderNo > 300000 and OrderNo < 400000
    taxstate = TaxState
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "ogsretotaller"
case OrderNo > 400000 and OrderNo < 500000
    taxstate=TaxState
    checkorder=replace(checkorder,"-",¬)
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "treeretotaller"
case OrderNo > 500000 and OrderNo < 600000
    taxstate = TaxState
    ;;arrayfilter checkorder, checkorder,¶, arrayinsert(extract(checkorder,¶,seq()),3,1,¬)
    ;;arrayfilter checkorder, checkorder,¶, arrayinsert(extract(checkorder,¶,seq()),5,1,¬)
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "bulbsretotaller"
case OrderNo > 600000 and OrderNo < 700000
    taxstate = TaxState
    checkorder=replace(checkorder,"–",¬)
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "mtretotaller"
endcase
call ".retotal"

___ ENDPROCEDURE .refigure _____________________________________________________

___ PROCEDURE .refiguretreeshipping ____________________________________________
local State
State=""
if memArray="Y"
    «$Shipping»=?(AdjTotal≤63, 10, AdjTotal*.16)
endif
if memArray="N"
    Case ShipCode="U"
        «$Shipping»=?(AdjTotal≤141, 22.5, AdjTotal*.16)
    Case ShipCode="X"
        «$Shipping»=?(AdjTotal≤125,25,AdjTotal*.2)
    DefaultCase
        «$Shipping»=0
    EndCase
endif
if Sta="ME"
«$Shipping»=10
endif
if ShipCode contains "T" or ShipCode contains "P"
«$Shipping»=0
endif
if Subtotal=0
«$Shipping»=0
endif
___ ENDPROCEDURE .refiguretreeshipping _________________________________________

___ PROCEDURE .refigureshipping ________________________________________________
global «$Sh», VZ, Vship, ssta, ppp
ssta=""
ppp=0
waswindow=info("windowname")
Vship=ShippingWt
VZ=Zip

if ShipCode contains "U" or ShipCode contains "X"
    Openfile "45shiplookup"
    Find ZipBegin<val(VZ) And val(VZ)≤ZipEnd
    Case Vship = 0
         «$Sh»=0
    Case Vship ≤ 2
        «$Sh»=«≤2»
    Case Vship ≤ 5
        «$Sh»=«≤5»
    Case Vship≤10
        «$Sh»=«≤10»
    Case Vship≤15
        «$Sh»=«≤15»
    Case Vship≤20
        «$Sh»=«≤20»
    Case Vship≤25
        «$Sh»=«≤25»
    Case Vship≤30
        «$Sh»=«≤30»
    Case Vship≤35
        «$Sh»=«≤35»
    Case Vship≤45
        «$Sh»=«≤45»
    Case Vship≤200 
        «$Sh»=Vship*«>45»
    Case Vship≤500
        «$Sh»=Vship*«≥200»
    Case Vship>500
        «$Sh»=Vship*«≥500»
    EndCase
    closewindow
    window waswindow
endif

if ShipCode="J"
    openfile "45depotshipping"
    Find ZipBegin<val(VZ) And val(VZ)≤ZipEnd
    «$Sh»=Vship*«per_lb»
    closewindow
    window waswindow
endif

if ShipCode="J"
    ssta=Depot
    openfile "45depotshipping"
    Find «depot_code»=ssta
    ppp=«price_per_pound»
    window waswindow
    «$Sh»=round(ShippingWt*ppp+.0001,.01)
    «$Sh»=?(«$Sh»<3.00,3.00,«$Sh»)
    
    if (Order contains "9276" or Order contains "9279")
        «$Sh»=«$Sh»+15.00
    endif
Endif

if ShipCode="H" and Depot contains "NOFA"
    «$Sh»=0
endif

«$Shipping»=«$Sh»
___ ENDPROCEDURE .refigureshipping _____________________________________________

___ PROCEDURE .retotal _________________________________________________________
case OrderNo > 700000 and OrderNo < 1000000
    ;; SEEDS
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=?(MemDisc>0, float(Subtotal*.01),0)
    AdjTotal=Subtotal-VolDisc-MemDisc
    
    if Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif
    
    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»+Donation+Membership
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    
case OrderNo > 300000 and OrderNo < 400000
    ;; OGS
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=?(MemDisc>0, float(Subtotal*.01),0)
    AdjTotal=Subtotal-VolDisc-MemDisc
    
    if Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif
    
    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»    
    GrTotal=«OrderTotal»+Donation
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    
case OrderNo > 600000 and OrderNo < 700000
    ;; POE
    VolDisc=float(«4SpareMoney»)*float(Discount)
    MemDisc=?(MemDisc>0, float(Subtotal*.01),0)
    AdjTotal=Subtotal-VolDisc-MemDisc
    
    if Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif
    
    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»+Donation
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    
case OrderNo > 400000 and OrderNo < 500000
    ;; TREES
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=?(MemDisc>0, float(Subtotal*.01),0)
    AdjTotal=Subtotal-VolDisc-MemDisc
    
    if Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        // enabled salestax recalculation for refiguring orders.
        call ".salestax"
    endif
    
    if «$Shipping»>22.50  //and MemDisc>0
        memArray="N"
        ;message "Please check shipping and adjust if needed."
        call ".refiguretreeshipping"
    endif
    
    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»+Donation+«Donation2»
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax
    «BalDue/Refund»=Paid-GrTotal

case OrderNo > 500000 and OrderNo < 600000
    ;; BULBS
    VolDisc=float(Subtotal)*float(Discount)
    MemDisc=?(MemDisc>0, float(Subtotal*.01),0)
    AdjTotal=Subtotal-VolDisc-MemDisc
    BulbsAdjTotal=BulbsSubtotal-(VolDisc+MemDisc)*divzero(BulbsSubtotal,Subtotal)
    if AdjTotal = 0
        BulbsAdjTotal = 0
    endif

    if ShipCode contains "T" or ShipCode contains "P"
        «$Shipping»=0
    else
        «$Shipping»=?(BulbsAdjTotal>100, BulbsAdjTotal*.12,12)
    
        If Sta contains "HI" or Sta contains "AK" or Sta contains "APO"
            «$Shipping»=?(BulbsAdjTotal>100, BulbsAdjTotal*.16,16)
        endif
    endif
    
    «$Shipping»=?(BulbsAdjTotal=0,0,«$Shipping»)
    
    if Taxable="Y"
        ;; disabling sales tax recalculation because they were overly simplistic. it works in paper collator, internet.inter, and tally
        ;call ".salestax"
    endif
    
    OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»
    GrTotal=«OrderTotal»
    RealTax=SalesTax
    Patronage=«OrderTotal»-RealTax

endcase
___ ENDPROCEDURE .retotal ______________________________________________________

___ PROCEDURE .staffrefigure ___________________________________________________
global checkorder
waswindow=info("windowname")
checkorder=""
checkorder=Order
case OrderNo > 700000 and OrderNo < 1000000
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "seedsretotaller"
    Subtotal=float(Subtotal*.5)
    Discount=0
    MemDisc=0
case OrderNo > 300000 and OrderNo < 400000
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "ogsstaffretotaller"
case OrderNo > 400000 and OrderNo < 500000
    checkorder=replace(checkorder,"-",¬)
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "treestaffretotaller"
case OrderNo > 500000 and OrderNo < 600000
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "bulbstaffretotaller"
case OrderNo > 600000 and OrderNo < 700000
    checkorder=replace(checkorder,"–",¬)
    openfile "OrderRetotaller"
    openfile "&@checkorder"
    call "mtstaffretotaller"
endcase
call ".retotal"
___ ENDPROCEDURE .staffrefigure ________________________________________________

___ PROCEDURE .yadayada ________________________________________________________
If info("trigger")="Button.SeeNote"
    comments=?(comments≠"","See note on order form."+¶+comments,"See note on order form.")
endif
___ ENDPROCEDURE .yadayada _____________________________________________________

___ PROCEDURE (Order Entry) ____________________________________________________

___ ENDPROCEDURE (Order Entry) _________________________________________________

___ PROCEDURE testzip __________________________________________________________
global Zip2

Zip2=Zip

displaydata Zip2

stop


global zipwindow
zipwindow="ZipCodeList:secret"
waswindow=info("windowname")
vzip=Zip
window zipwindow
select ZipCode=vzip
arrayselectedbuild addressArray,¶, "ZipCodeList", City+¬+State
selectall
window waswindow
case arraysize(addressArray,¶)=0
    message "ZipCode not found, try again"
case arraysize(addressArray,¶)=1
    City=extract(addressArray,¬,1)
    St=extract(addressArray,¬,2)
case arraysize(addressArray,¶)>1
    superchoicedialog addressArray, rayi, {height=400 width=500 font=Helvetica caption="Click on one and then hit OK or New for new entry" 
        captionfont=Times captionsize=12 captioncolor=red size=14 buttons="OK:100;Cancel:100"}
    if info("dialogtrigger") contains "OK"
        City=extract(rayi,¬,1)
        St=extract(rayi,¬,2)
    endif
    if info("dialogtrigger") contains "Cancel"
        stop
    endif
endcase
___ ENDPROCEDURE testzip _______________________________________________________

___ PROCEDURE nextorder/√ ______________________________________________________
if info("selected")<info("records")
    SelectAll
endif

getscrap "Next!"

if length(str(clipboard()))=6
    case str(clipboard()) beginswith "3"
        goform "ogsinput"
    case str(clipboard()) beginswith "4"
        goform "treesinput"
    case str(clipboard()) beginswith "5"
        goform "bulbsinput"
    case str(clipboard()) beginswith "6"
        goform "mtinput"
    case str(clipboard()) beginswith "7"
        goform "seedsinput"
    endcase
    oono=0
    find OrderNo=val(clipboard())
    if info("found")=0
        beep
        message "Sorry! Not here"
        stop
    endif
    field «C#»
    if «C#»>0
        field Sub1
        stop
    endif
    editcellstop
endif

Numb=val(clipboard())
case info("windowname")=newyear+"orders:seedsinput"
    Numb=?(Numb<100000, Numb+700000,Numb)
case info("windowname")=newyear+"orders:ogsinput"
    Numb=?(Numb<100000, Numb+300000,Numb)
case info("windowname")=newyear+"orders:treesinput"
    Numb=?(Numb<100000, Numb+400000,Numb)
case info("windowname")=newyear+"orders:bulbsinput"
    Numb=?(Numb<100000, Numb+500000,Numb)
case info("windowname")=newyear+"orders:mtinput"
    Numb=?(Numb<100000, Numb+600000,Numb)
endcase
find OrderNo=Numb
field «C#»

if «C#»>0
    stop
else
    editcellstop
endif
___ ENDPROCEDURE nextorder/√ ___________________________________________________

___ PROCEDURE export/π _________________________________________________________
Window newyear+"orders:addresschecker"
YesNo "do you want to export?"
If clipboard()="No"
    LastRecord
    If info("Summary")≠0
        DeleteRecord
    Endif
    SelectAll
    Save
    LastRecord
    CloseWindow
    Stop
EndIf

LastRecord
If info("summary")≠0
    DeleteRecord
Endif
firstrecord
field OrderBatched
Formulafill today()

loop
    if ShipCode="C" and «$Shipping»=0
        CreditCard=CreditCard
    else
        CreditCard= ?(length(CreditCard)<15, CreditCard,"XXXX"+CreditCard[-4,-1])
    endif
    downrecord
until stopped

Flag=""
OpenFile newyear+"orders/export"
OpenFile "&&"+newyear+"orders"
FirstRecord
;If «C#»=0
;DeleteRecord
;EndIf
field OrderNo
SaveACopyAs Dialog
Save
CloseFile
Window newyear+"orders:addresschecker"
selectall
CloseWindow
find OrderNo=Numb
Field OrderNo
Save

___ ENDPROCEDURE export/π ______________________________________________________

___ PROCEDURE sortup/0 _________________________________________________________
synchronize
field OrderNo
sortup
___ ENDPROCEDURE sortup/0 ______________________________________________________

___ PROCEDURE next/1 ___________________________________________________________

//Is this an internetorder?
call .IsInternetOrder

global WinCheck, intOrder1, fromBranch,restartRun,formCheck


intOrder1=int(OrderNo)
fromBranch=""
WinCheck=info("windows")
formCheck=info("formname")

//Get User to the right forms
if WinCheck notcontains "input" or formCheck notcontains "input"
    message "This Must be run from an 'order Input' form"
    stop
endif


//___Check Selection/Entered___//
if val(EntryDate)>0
    yesno "This Order Has already been entered, would you like to select Orders that have not been entered?"
    if clipboard()="Yes"
        select OrderNo>0
        //check about this selection criteria with Bria
        selectwithin Sequence≠0
        selectwithin val(EntryDate)=0
        firstrecord
        restartRun=-1
        Goto CheckForm
    else
        stop
    endif
endif

//_____CheckForm_________//


restartRun=0

CheckForm:

//___duplicated for checkform looping___//
if restartRun=-1
    intOrder1=int(OrderNo)
    fromBranch=""
    WinCheck=info("windows")
endif
//____________________________________//

case intOrder1 ≥ 700000
    fromBranch="Seeds"
        //__opens or goes to open input window__//
        if WinCheck notcontains "seedsinput"
            setwindow 336,132,643,844,""
            openform "seedsinput"
        endif

        if info("formname") notcontains "seedsinput"
            goform "seedsinput"
        endif  
case intOrder1 ≥ 600000 and intOrder1 < 700000
    fromBranch="OGS;POE"
        if WinCheck notcontains "mtinput"
            setwindow 336,132,588,977,""
            openform "mtinput"
         endif

        if info("formname") notcontains "mtinput"
            goform "mtinput"
        endif 
case intOrder1 ≥ 500000 and intOrder1 < 600000
    fromBranch="OGS;Bulbs"
        if WinCheck notcontains "bulbsinput"
            setwindow 356,71,521,856,""
            openform "bulbsinput"
        endif

        if info("formname") notcontains "bulbsinput"
            goform "bulbsinput"
        endif 
case intOrder1 ≥ 400000 and intOrder1 < 500000
    fromBranch="Trees"
        if WinCheck notcontains "treesinput"
            setwindow 339,79,591,779,""
            openform "treesinput"
        endif

        if info("formname") notcontains "treesinput"
            goform "treesinput"
        endif 
case intOrder1 ≥ 300000 and intOrder1 <400000
    fromBranch="OGS"
        if WinCheck notcontains "ogsinput"
            setwindow 420,77,521,856,""
            openform "ogsinput"
        endif

        if info("formname") notcontains "ogsinput"
            goform "ogsinput"
        endif 
endcase
//____EndCheckForm______//




groupArray=?(Group≠"", Group, Con)
;rayc=«C#Text»

//__apply discount__//
if OrderNo=int(OrderNo)
    rayg=?(MemDisc=round(Subtotal*.01,.01),"Y","")
endif

//__getPool__//
if info("formname")="treesinput"
    rayi=Pool
endif

//downrecords if not starting over

ono=OrderNo
if restartRun=0
    downrecord
endif
oono=0

//___Not Sure why this is here__ ASK KATE AND BRIA IF THEY NEED CMD-1 TO WORK WITH THIS//
if info("formname")="addresschecker" 
    stop
endif

field «C#»
if OrderNo≠int(OrderNo)
    Group=groupArray
    ;«C#Text»=rayc
    MemDisc=?(rayg="Y",.01,MemDisc)
    if info("formname")="treesinput"
        Pool=rayi
    endif
    field Con
    ;call ".dropship"
endif

if str(OrderNo) contains "."
    EntryDate = today()
    stop
else
    EntryDate=today()
endif

;; just internet orders

if OrderNo=int(OrderNo)
    if (OrderNo >= 320000 and OrderNo < 400000)
    or (OrderNo >= 420000 and OrderNo < 500000)
    or (OrderNo >= 520000 and OrderNo < 600000)
    or (OrderNo >= 620000 and OrderNo < 700000)
    or (OrderNo >= 710000 and OrderNo < 1000000)
            ;call "updatemail/7"
            call "NewSearch/ß"
    endif
endif

if restartRun=-1
downrecord
endif

///________testspace____Delete me______////



___ ENDPROCEDURE next/1 ________________________________________________________

___ PROCEDURE sameship/2 _______________________________________________________
SAd=MAd
Cit=City
Sta=St
Z=Zip
Call ".Z"
if ShipCode contains "P" or ShipCode contains "T"
    TaxState="ME"
    TaxRate=.055
else
    TaxState=Sta
TaxRate=lookup("ZipCodeList","ZipCode", «Z»,"TaxRate",0,0)
StateRate=lookup("ZipCodeList","ZipCode", «Z»,"StateRate",0,0)
CountyRate=lookup("ZipCodeList","ZipCode", «Z»,"CountyRate",0,0)
CityRate=lookup("ZipCodeList","ZipCode", «Z»,"CityRate",0,0)
SpecialRate=lookup("ZipCodeList","ZipCode", «Z»,"SpecialRate",0,0)
endif
___ ENDPROCEDURE sameship/2 ____________________________________________________

___ PROCEDURE copycity/3 _______________________________________________________
Cit=City
Sta=St
Z=Zip
Field SAd
EditCell
Call ".Z"

___ ENDPROCEDURE copycity/3 ____________________________________________________

___ PROCEDURE justlocate/4 _____________________________________________________
getscrap "What order? (Please use all 6 digits)"
ono=val(clipboard())
find OrderNo=ono
case ono>700000 and ono<1000000
    GoForm "seedsinput"
case ono>300000 and ono<400000
    goform "ogsinput"
case ono>400000 and ono<500000
    goform "treesinput"
case ono>500000 and ono<600000
    goform "bulbsinput"
case ono>600000 and ono<700000
    goform "mtinput"
endcase

___ ENDPROCEDURE justlocate/4 __________________________________________________

___ PROCEDURE writeemail/5 _____________________________________________________
local mailaddress, mailcopies, mailheader, messageBody
applescript |||
tell application "Mail"
activate
end tell
tell application "Panorama"
activate
end tell |||
opensecret "email messages"
mailaddress=""
mailcopies=""
mailheader=""
messageBody=""
setwindowrectangle
rectanglesize(59, 107, 689, 663), ""
OpenForm "Emailform"
if Email≠''
mailaddress=Email
mailheader= "Your Fedco Order "+str(OrderNo)
superobject "emailaddress", "open", "inserttext", mailaddress
activesuperobject "close"
endif
superobject "emailsubject", "open", "inserttext", mailheader
activesuperobject "close"

___ ENDPROCEDURE writeemail/5 __________________________________________________

___ PROCEDURE same mail/6 ______________________________________________________
MAd=SAd
City=Cit
St=Sta
Zip=Z
___ ENDPROCEDURE same mail/6 ___________________________________________________

___ PROCEDURE updatemail/7 _____________________________________________________
global selectedAddressArray, chosenAddress, chosenAddress1, FindWindow, WinNumber


waswindow=info("windowname")

;call ".dropship"
if info("formname")="treesinput" and (OrderNo<410000 or OrderNo>420000)
    call ".pool"
endif
EntryDate=today()
ono=OrderNo
vzip=Zip
vd=«C#»
conArray=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1]
place=MAd["0-9",-1][1,2]
field «1stPayment»
;editcell

WinNumber=arraysearch(info("windows"), thisFYear+" mailing list", 1,¶)
if WinNumber=0
    openfile newyear+" mailing list"
endif
window newyear+" mailing list"
if vd>0
    find «C#»=vd
    if info("found")=-1
        YesNo "Enter this one?"
        if clipboard() contains "Yes"
            call "enter/e"
            stop
        endif
    endif

    if info("found")=0
        goto newzip
    endif
    window newyear+" mailing list"
endif

//*********************************************************//
//Searches by Zip
newzip:

find Zip=vzip

//Zip that's not in Mailaing List
if info("found")=0
    find Con contains extract(conArray," ",1)and Con contains extract(conArray," ",2)
    if info("Found")=-1
        goto newentry
    else
        addmail:

        insertrecord
        Con=grabdata(newyear+"orders",Con) 
        Group=grabdata(newyear+"orders",Group) 
        MAd=grabdata(newyear+"orders",MAd) 
        City=grabdata(newyear+"orders",City) 
        St=grabdata(newyear+"orders",St)
        Zip=grabdata(newyear+"orders",Zip) 
        adc=lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        SAd=grabdata(newyear+"orders",SAd) 
        Cit=grabdata(newyear+"orders",Cit) 
        Sta=grabdata(newyear+"orders",Sta) 
        Z=grabdata(newyear+"orders",Z) 
        phone=grabdata(newyear+"orders",Telephone) 
        email=grabdata(newyear+"orders", Email)
        adc=lookup("fcmadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        if ono>300000 and ono<400000
            call "ogsity/ø"
        endif
        if ono>500000 and ono<600000
            call "bulbous/∫"
        endif
        if ono>400000 and ono<500000
            call "treed/†"
        endif
        if ono>700000 and ono < 1000000
            call "seedy/ß"
        endif
        if ono>600000 and ono<700000
            call "moosed/µ"
        endif
        window waswindow
    endif
endif


//Zip is in Mailing List
if info("found")=-1
    if place ≠ ""
        find MAd contains place And Zip=vzip
        if info("found")=0
            find Zip=vzip and Con contains extract(conArray," ",2)
            if info("found")=-1
                goto newentry
            endif
                insertrecord
                Con=grabdata(newyear+"orders",Con) 
                Group=grabdata(newyear+"orders",Group) 
                MAd=grabdata(newyear+"orders",MAd) 
                City=grabdata(newyear+"orders",City) 
                St=grabdata(newyear+"orders",St)
                Zip=grabdata(newyear+"orders",Zip) 
                adc=lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
                SAd=grabdata(newyear+"orders",SAd) 
                Cit=grabdata(newyear+"orders",Cit) 
                Sta=grabdata(newyear+"orders",Sta) 
                Z=grabdata(newyear+"orders",Z) 
                phone=grabdata(newyear+"orders",Telephone) 
                email=grabdata(newyear+"orders", Email)
                if ono>300000 and ono<400000
                call "ogsity/ø"
                endif
                if ono>500000 and ono<600000
                call "bulbous/∫"
                endif
                if ono>400000 and ono<500000
                call "treed/†"
                endif
            if ono>700000 and ono < 1000000
                call "seedy/ß"
                endif
                if ono>600000 and ono<700000
                call "moosed/µ"
                endif
        endif
        if info("found")=-1
           newentry:
            select MAd contains place And Zip=vzip
            selectedAddressArray=""
            arrayselectedbuild selectedAddressArray, ¶,"",str(«C#»)+¬+«Con»+¬+«Group»+¬+«MAd»+¬+«City»+¬+«St»
            selectall
            find Zip=vzip
            superchoicedialog selectedAddressArray, chosenAddress, 
                {title="Choose the Correct Customer/Address" caption="Click -Other Search- to Search by something else" buttons="ok;other search;cancel"}
            if info("dialogtrigger") contains "ok"
                find exportline() contains chosenAddress
                if info("found")=-1
                    message "Found!"
                    call "enter/e"
                else
                    message "error, repeating search..."
                    farcall "45orders","NewSearch/`"
                endif
            endif
            if info("dialogtrigger") contains "search"
                farcall "45orders","NewSearch/`"
            endif
            if info("dialogtrigger") contains "cancel"
            stop
            endif
             goto addmail
        endif
    endif
endif
___ ENDPROCEDURE updatemail/7 __________________________________________________

___ PROCEDURE ordersonly/8 _____________________________________________________
GetScrap "Which orders"
case clipboard()="seeds"
    Select OrderNo ≥ 700000 and OrderNo < 1000000
Case clipboard()="bulbs"
    Select OrderNo ≥ 500000 And OrderNo < 600000
Case clipboard()="ogs"
    Select OrderNo ≥ 300000 And OrderNo < 400000
Case clipboard()="moose"
    Select OrderNo ≥ 600000 And OrderNo < 700000
Case clipboard()="trees"
    Select OrderNo ≥ 400000 And OrderNo < 500000
Case clipboard()=""
    SelectAll
EndCase
SelectWithin OrderNo=int(OrderNo)
___ ENDPROCEDURE ordersonly/8 __________________________________________________

___ PROCEDURE daily/9 __________________________________________________________
waswindow= info("windowname")
Numb=0
Save
GetText "enter first order #", Numb

Numb=val(Numb)

case info("windowname")=newyear+"orders:seedsinput"
    Numb=?(Numb<100000, Numb+700000,Numb)
case info("windowname")=newyear+"orders:ogsinput"
    Numb=?(Numb<100000, Numb+300000,Numb)
case info("windowname")=newyear+"orders:treesinput"
    Numb=?(Numb<100000, Numb+400000,Numb)
case info("windowname")=newyear+"orders:bulbsinput"
    Numb=?(Numb<100000, Numb+500000,Numb)
case info("windowname")=newyear+"orders:mtinput"
    Numb=?(Numb<100000, Numb+600000,Numb)
endcase

Select OrderNo ≥ Numb
GetText "enter last order #", Numb
Numb=val(Numb)
case info("windowname")=newyear+"orders:seedsinput"
    Numb=?(Numb<100000, Numb+700000,Numb)
case info("windowname")=newyear+"orders:ogsinput"
    Numb=?(Numb<100000, Numb+300000,Numb)
case info("windowname")=newyear+"orders:treesinput"
    Numb=?(Numb<100000, Numb+400000,Numb)
case info("windowname")=newyear+"orders:bulbsinput"
    Numb=?(Numb<100000, Numb+500000,Numb)
case info("windowname")=newyear+"orders:mtinput"
    Numb=?(Numb<100000, Numb+600000,Numb)
endcase
SelectWithin OrderNo < Numb+1
WindowBox "50 50 460 677"
OpenForm "addresschecker"
FirstRecord


___ ENDPROCEDURE daily/9 _______________________________________________________

___ PROCEDURE chargecards/® ____________________________________________________
global orderwindow
orderwindow=info("windowname")
getscrap "first order to charge (6 digits)"
select OrderNo≥val(clipboard())
getscrap "last order to charge (6 digits)"
selectwithin OrderNo≤val(clipboard())
selectwithin CreditCard≠""
field Paid
formulafill GrTotal
field «1stPayment»
formulafill GrTotal
field «BalDue/Refund»
formulafill GrTotal-Paid
export "exportcharges", ¬+str(«OrderNo»)+¬+pattern(Zip,"#####")+¬+Email
    +¬+pattern(GrTotal,"#.##")+¬+¬+¬+CreditCard+¬+ExDate+¶
openfile "problem orders"
windowtoback "problem orders"
openfile "credit charges"
openfile "&exportcharges"
call "charge"    
___ ENDPROCEDURE chargecards/® _________________________________________________

___ PROCEDURE checkdups/ç ______________________________________________________
waswindow=info("windowname")
getscrap "which division?"
case clipboard()="seeds"
    select OrderNo > 700000 and OrderNo < 1000000
case clipboard()="ogs"
    select OrderNo > 300000 and OrderNo < 400000
case clipboard()="trees"
    select OrderNo > 400000 and OrderNo < 500000
case clipboard()="bulbs"
    select OrderNo > 500000 and OrderNo < 600000
case clipboard()="mt"
    select OrderNo > 600000 and OrderNo < 700000
endcase

selectwithin Zip≠0

local raya, rayb, vzip,vtot,ono,vcon
rayb=""
hide
field Zip
sortup
field Subtotal
sortupwithin
firstrecord

loop
    ono=OrderNo
    vzip=Zip
    vtot=Subtotal
    vcon=Con
    downrecord
    repeatloopif OrderNo≠int(OrderNo)
    stoploopif info("eof")
    if Zip=vzip and Subtotal=vtot and Con=vcon
        raya=str(ono)+" is same as "+str(OrderNo)
        rayb=rayb+¶+raya
    endif
until info("stopped")
show
selectall
field OrderNo
sortup 
firstrecord
if rayb=""
    message "no dups"
    stop
endif
goform "duplicates"


___ ENDPROCEDURE checkdups/ç ___________________________________________________

___ PROCEDURE hide/h ___________________________________________________________
window "Hide This Window"
window newyear+" mailing list"
saveall
___ ENDPROCEDURE hide/h ________________________________________________________

___ PROCEDURE email/m __________________________________________________________
waswindow=info("windowname")
ID=«C#»
window newyear+" mailing list:secret"
If «C#»≠ID
    Find «C#»=ID
EndIf
SAd=grabdata(newyear+"orders",SAd)
Cit=grabdata(newyear+"orders",Cit)
Sta=grabdata(newyear+"orders",Sta)
Z=grabdata(newyear+"orders",Z)
email=grabdata(newyear+"orders",Email)
phone=grabdata(newyear+"orders",Telephone)
window waswindow
;if OrderNo=oono
    ;stop 
;endif
;if info("formname")="ogsinput"
    ;field «1stPayment»
    ;editcell
    ;stop
;else
    ;if info("formname")="treesinput"
        ;call ".pool"
    ;endif
    ;Field Sub1
    ;editcell
    ;field Sub2
    ;editcell
    ;Field «1stPayment»
    ;stop
;EndIf
___ ENDPROCEDURE email/m _______________________________________________________

___ PROCEDURE (print) __________________________________________________________

___ ENDPROCEDURE (print) _______________________________________________________

___ PROCEDURE treenotes ________________________________________________________
select OrderNo>400000 and OrderNo<500000
selectwithin CustomerComments≠"" or OrderComments≠""
openform "treenotes"
Print ""
closewindow
goform "treesinput"
___ ENDPROCEDURE treenotes _____________________________________________________

___ PROCEDURE printpickuplist __________________________________________________
waswindow=info("windowname")
local firstorder, lastorder, firstintorder, lastintorder
getscrap "first seed order, use 6 digits"
firstorder=val(clipboard())
getscrap "last seed order, use 6 digits"
lastorder=val(clipboard())
select OrderNo≥firstorder And OrderNo≤lastorder
getscrap "first internet seed order, use 6 digits"
firstintorder=val(clipboard())
getscrap "last internet seed order, use 6 digits"
lastintorder=val(clipboard())
selectadditional OrderNo≥firstintorder And OrderNo≤lastintorder
getscrap "first ogs order, use 6 digits"
firstorder=val(clipboard())
getscrap "last ogs order, use 6 digits"
lastorder=val(clipboard())
selectadditional OrderNo≥firstorder And OrderNo≤lastorder
getscrap "first internet ogs order, use 6 digits"
firstintorder= val(clipboard())
getscrap "last internet ogs order, use 6 digits"
lastintorder=val(clipboard())
selectadditional OrderNo≥firstintorder And OrderNo≤lastintorder
selectwithin ShipCode="P"
selectwithin OrderNo=int(OrderNo)
openfile newyear+"postcardfile"
openfile "&&"+newyear+"orders"
stop
call "printpostcards"
___ ENDPROCEDURE printpickuplist _______________________________________________

___ PROCEDURE postcards in march _______________________________________________
global firstorder, lastorder, cust
getscrap "first seed order, use 5 digits"
firstorder=val(clipboard())
getscrap "last seed order, use 5 digits"
lastorder=val(clipboard())
select OrderNo≥firstorder And OrderNo≤lastorder
selectwithin ShipCode="P"
selectwithin OrderNo=int(OrderNo)
firstrecord

loop 
    ono=OrderNo
    cust=«C#»
    selectadditional «C#»=cust and (OrderNo>300000 and OrderNo<400000) and ShipCode="P"
    find OrderNo=ono
    downrecord
    stoploopif OrderNo>300000
until info("stopped")

openform "pickups"
print ""
message "Change paper"
closewindow
openform "postcards:seeds & comb"
selectwithin OrderNo<300000

loop
    printonerecord
    downrecord
until info("stopped")

closewindow
selectall
save
___ ENDPROCEDURE postcards in march ____________________________________________

___ PROCEDURE print MT pickups _________________________________________________
select ShipCode="P" and OrderNo>600000 and OrderNo<700000
selectwithin OrderNo=int(OrderNo)
field OrderNo
sortup

 
___ ENDPROCEDURE print MT pickups ______________________________________________

___ PROCEDURE print moose cards ________________________________________________
local cust
select OrderNo>600000 and OrderNo<700000
selectwithin ShipCode="P"
selectwithin OrderNo=int(OrderNo)
openfile newyear+"postcardfile"
openfile "&&"+newyear+"orders"
field «C#»
groupup
field ShippingWt
total
field Con
propagate
field Group
emptyfill " "
propagate
field MAd
propagate
field City
propagate
field St
propagate
field Zip
propagate
field «C#Text»
propagate
lastrecord
deleterecord
selectsummaries
stop
openform "postcards:mt"
print dialog
closewindow
openform "mtpickups"
print dialog
closewindow
window waswindow
selectall

___ ENDPROCEDURE print moose cards _____________________________________________

___ PROCEDURE (extras) _________________________________________________________

___ ENDPROCEDURE (extras) ______________________________________________________

___ PROCEDURE changepd _________________________________________________________
Paid=0
«1stPayment»=0
AddPay1=0
«BalDue/Refund»=Paid-GrTotal
___ ENDPROCEDURE changepd ______________________________________________________

___ PROCEDURE delete card ______________________________________________________
field CreditCard

formulafill ?(CreditCard="" or CreditCard contains "XXXX",CreditCard,"XXXX"+CreditCard[-4,-1])

___ ENDPROCEDURE delete card ___________________________________________________

___ PROCEDURE moosetodate ______________________________________________________
local newarray, exported
newarray=""
exported=""
selectwithin Order <> ""
firstrecord
loop

arrayfilter Order, newarray, ¶, arraydelete(extract(Order,¶,seq()),1,1,¬)

newarray=replace(newarray, "–",¬)

arrayfilter newarray, newarray, ¶, arraydelete(extract(newarray, ¶, seq()), 4,2,¬)
arrayfilter newarray, newarray, ¶, arraydelete(extract(newarray, ¶, seq()), 7,1,¬)
arrayfilter newarray, newarray,¶, ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+Sub1+¬+Sub2
    +¬+datepattern(OrderPlaced,"mm/dd/yyyy")+¬+datepattern(EntryDate,"mm/dd/yyyy")+¬+import()+¬+
    arraystrip(stripchar(CustomerComments,"!ÿ "),¬)+¬+arraystrip(stripchar(OrderComments,"!ÿ "),¬)+¬+
    arraystrip(stripchar(Con,"!ÿ "),¬)+¬+Telephone+¬+arraystrip(arraystrip(Email,¶),¬)
exported=exported+newarray+¶
newarray=""
downrecord
until info("stopped")
openfile "mooseboughttodate"
openfile "&@exported"
;;call "exporttext/4"
call "distribution"
___ ENDPROCEDURE moosetodate ___________________________________________________

___ PROCEDURE treesboughttodate ________________________________________________
local newarray, exported
newarray=""
exported=""

//2015 -added "xxtreescurrent" as a delinked xxtrees copy for Jen to view all orders with customer contact info
openfile "45treescurrent"
openfile "&&45trees"
OpenForm "facilitator view"
saveacopyas "45treescurrent "+datepattern(today(), "mm/dd/yy")
closefile

select OrderNo>400000 and OrderNo<500000
field OrderNo
sortup
selectwithin ShipCode notcontains "D"
selectwithin ShipCode notcontains "C"
firstrecord

loop
newarray=Order
newarray=replace(newarray,"-A","1")
newarray=replace(newarray, "-B","2")
newarray=replace(newarray, "-C","3")
newarray=replace(newarray,"-D","4")
newarray=replace(newarray,"-E","5")
newarray=replace(newarray,"-F","6")
newarray=replace(newarray,"-G","7")
newarray=replace(newarray, "-H","8")
newarray=replace(newarray, "-J","9")


arrayfilter newarray, newarray, ¶, arraydelete(extract(newarray,¶,seq()),1,1,¬)
;arrayfilter newarray, newarray, ¶, arraydelete(extract(newarray, ¶, seq()), 4,2,¬)
newarray=arraystrip(newarray,¬)

arrayfilter newarray, newarray,¶, ShipCode+¬+str(Pool)+¬+datepattern(OrderPlaced, "mm/dd/yy")+¬+Sub1+¬+Sub2+¬+str(OrderNo)+¬+str(Sequence)+¬+import()
exported=exported+newarray+¶
newarray=""

downrecord
until info("stopped")
openfile "treesboughttodate"
openfile "&@exported"
call ".distribution"
___ ENDPROCEDURE treesboughttodate _____________________________________________

___ PROCEDURE textexport _______________________________________________________
local importarray
ArraySelectedBuild importarray,¶, info("DatabaseName"), str(«C#»)+","+str(OrderNo)+","+
        str(OrderNo)+": "+Con
        +","+?(Group≠"",Group,Con)+","+
        ?(SAd contains "-", SAd[1,"-"][1,-2]+","+SAd["-",-1][2,-1],SAd+","+".")
        +","+Cit+","+Sta+","+pattern(Z,"#####")+","+
        ?(phone≠"",?(length(phone)=8, 
        "207-"+phone, ?(phone contains "/",phone[1,"/"][1,-2],phone)),".")+","+"Y"
importarray="CustID"+","+"OrderNo"+","+"Con"+","+"Group"+","+"SAd1"+","+"SAd2"+","+
    "City"+","+"State"+","+"Zip"+","+"Phone"+","+"Res"+¶+importarray
Select OrderNo=300001    
export "import.txt", importarray   
selectall 
___ ENDPROCEDURE textexport ____________________________________________________

___ PROCEDURE ShipLookup _______________________________________________________
Numb=int(OrderNo)

case ShipCode="U" OR ShipCode="X" or ShipCode="C"
OpenFile newyear+"shipping"
selectall
Synchronize
select «O#»=Numb OR «C#»= str(Numb)

case ShipCode="P" OR ShipCode="T"
YesNo "This order is a pickup. Do you still want to check shipping?"
    if clipboard()="Yes"
    OpenFile newyear+"shipping"
    selectall
Synchronize
    select «O#»=Numb OR «C#»= str(Numb)
    else
    stop
    endif

case ShipCode="H"
YesNo "This order is being held for some reason. Do you still want to check shipping?"
    if clipboard()="Yes"
    OpenFile newyear+"shipping"
    selectall
Synchronize
    select «O#»=Numb OR «C#»= str(Numb)
    else
    stop
    endif
endcase
if   info("Selected") < info("Records") 
stop
else
message "nothing found"
endif
___ ENDPROCEDURE ShipLookup ____________________________________________________

___ PROCEDURE ChangeCard _______________________________________________________
GetScrap "New card #, 16 digits."
if val(clipboard())>0 AND length(clipboard())=16
CreditCard=clipboard()
endif
if length(clipboard())≠16 AND length(clipboard())>0 AND clipboard()[1,1]≠"3"
Message "You've got the wrong number of digits. Try again."
stop
endif

GetScrap "NewDate, 4 digits."
if val(clipboard())>0 AND length(clipboard())=4
ExDate=clipboard()
endif
if length(clipboard())≠4 AND length(clipboard())>0
Message "You've got the wrong number of digits or format. Try again."
stop
endif

___ ENDPROCEDURE ChangeCard ____________________________________________________

___ PROCEDURE force unlock _____________________________________________________
forceunlockrecord
___ ENDPROCEDURE force unlock __________________________________________________

___ PROCEDURE forcesynchronize _________________________________________________
forcesynchronize
call "sortup/0"
___ ENDPROCEDURE forcesynchronize ______________________________________________

___ PROCEDURE export Canada ____________________________________________________
global ohcanada
ohcanada=""
ohcanada=Order
arrayfilter ohcanada, ohcanada,¶, arraydelete(extract(ohcanada,¶,seq()),1,1,¬)
arrayfilter ohcanada, ohcanada,¶, arraydelete(extract(ohcanada,¶,seq()), 4,3,¬)
openfile "Canadian orders"
openfile "+@ohcanada"
___ ENDPROCEDURE export Canada _________________________________________________

___ PROCEDURE addpay ___________________________________________________________
 local addpay, vhow,vorderref,vorderpay
 
 case info("trigger") = "Button.Undo AddPay1"
    Paid=Paid-AddPay1
    «BalDue/Refund»=«BalDue/Refund»-AddPay1
    AddPay1=0
    DatePay1=0
    MethodPay1=""
    stop 
case info("trigger") = "Button.Undo AddPay2"
    Paid=Paid-AddPay2
    «BalDue/Refund»=«BalDue/Refund»-AddPay2
    AddPay2=0
    DatePay2=0
    MethodPay2=""
    stop 
case info("trigger") = "Button.Undo AddPay3"
    Paid=Paid-AddPay3
    «BalDue/Refund»=«BalDue/Refund»-AddPay3
    AddPay3=0
    DatePay3=0
    MethodPay3=""
    stop 
case info("trigger") = "Button.Undo AddPay4"
    Paid=Paid-AddPay4
    «BalDue/Refund»=«BalDue/Refund»-AddPay4
    AddPay4=0
    DatePay4=0
    MethodPay4=""
    stop 
case info("trigger") = "Button.Undo AddPay5"
    Paid=Paid-AddPay5
    «BalDue/Refund»=«BalDue/Refund»-AddPay5
    AddPay5=0
    DatePay5=0
    MethodPay5=""
    stop 
case info("trigger") = "Button.Undo AddPay6"
    Paid=Paid-AddPay6
    «BalDue/Refund»=«BalDue/Refund»-AddPay6
    AddPay6=0
    DatePay6=0
    MethodPay6=""
    stop 
 endcase

 GetScrap  "What's the additional payment?"
 addpay=val(clipboard())
 getscrap "How was it paid? (cc, gc, ch, cash, tr, mc, ogc)"
 case (clipboard() contains "cc" or clipboard() contains "cred")
    if «BalDue/Refund» > 0 and today()-OrderPlaced≥180 ;; refund
        message "transaction too old: issue a check"
        stop
    else ;; balance due
        if «BalDue/Refund» > 0 and today()-OrderPlaced≥90
            message "transaction too old: can't rebill"
            stop
        endif
    endif
    vhow="Credit_Card" ;; only assign this vhow if it makes it past the date checks
 case (clipboard() contains "gc" or clipboard() contains "gift")
    vhow="Gift_Certificate"
 case (clipboard() contains "ch" or clipboard() contains "√")
    vhow="Check"
 case (clipboard() contains "cash" or clipboard() contains "$")
     vhow="Cash"
 case (clipboard() contains "mc")
     vhow="MOFGA_Certificate"
 case (clipboard() contains "ogc")
     vhow="Old_Gift_Certificate"
 endcase
 
 case clipboard() contains "tr"
    vhow="Transfer"
    vorderref=str(OrderNo)
    gettext "To which order are you transferring this payment?",vorderpay
    «Notes2»=«Notes2»+¶+"Payment transferred to order "+vorderpay
 endcase
 
 
;; Paid= Paid+addpay
 
 ;;«BalDue/Refund»=«BalDue/Refund»+addpay
 
 If «AddPay1»=0
    «AddPay1»= addpay
    «DatePay1»=today()
    «MethodPay1»=vhow
else
    if «AddPay2»=0
        «AddPay2»= addpay
        «DatePay2»=today()
        «MethodPay2»=vhow
    else
        if «AddPay3»=0
            «AddPay3»=addpay
            «DatePay3»=today()
            «MethodPay3»=vhow
        else
            if «AddPay4»=0
            «AddPay4»=addpay
            «DatePay4»=today()
            «MethodPay4»=vhow
            else
                if «AddPay5»=0
                «AddPay5»=addpay
                «DatePay5»=today()
                «MethodPay5»=vhow
                else
                    if «AddPay6»=0
                    «AddPay6»=addpay
                    «DatePay6»=today()
                    «MethodPay6»=vhow
                    else 
                    message "This order is crazy. Please consolidate additional payments."
                    stop
                    endif
                endif
           endif
        endif
    endif
endif


Paid = «1stPayment» + AddPay1 + AddPay2 + AddPay3 + AddPay4 + AddPay5 + AddPay6
«BalDue/Refund»=Paid-GrTotal
RealTax=?(info("formname")="seedsinput" or info("formname")="ogsinput" or info("formname")="mtinput",SalesTax*(1-Discount),SalesTax)
Patronage=OrderTotal-RealTax 

if vhow="Transfer"
    find OrderNo=val(vorderpay)
    if info("found")
        addpay=-addpay
        «Notes2»=«Notes2»+¶+"Payment transferred from order "+str(vorderref)
        ;; Paid= Paid+addpay
 
        ;; «BalDue/Refund»=«BalDue/Refund»+addpay
        If «AddPay1»=0
            «AddPay1»= addpay
            «DatePay1»=today()
            «MethodPay1»=vhow
        else
            if «AddPay2»=0
                «AddPay2»= addpay
                «DatePay2»=today()
                «MethodPay2»=vhow
            else
                if «AddPay3»=0
                    «AddPay3»=addpay
                    «DatePay3»=today()
                    «MethodPay3»=vhow
                else
                    if «AddPay4»=0
                        «AddPay4»=addpay
                        «DatePay4»=today()
                        «MethodPay4»=vhow
                    else
                        if «AddPay5»=0
                            «AddPay5»=addpay
                            «DatePay5»=today()
                            «MethodPay5»=vhow
                        else
                            if «AddPay6»=0
                                «AddPay6»=addpay
                                «DatePay6»=today()
                                «MethodPay6»=vhow
                            else 
                                message "This order is crazy. Please consolidate additional payments."
                                stop
                            endif
                        endif
                    endif
                endif
            endif
        endif
        yesno
        "Does this look right?"
        if clipboard() contains "n"
            stop
        endif
        Paid = «1stPayment» + AddPay1 + AddPay2 + AddPay3 + AddPay4 + AddPay5 + AddPay6
        «BalDue/Refund»=Paid-GrTotal
        RealTax=?(info("formname")="seedsinput" or info("formname")="ogsinput" or info("formname")="mtinput",SalesTax*(1-Discount),SalesTax)
        Patronage=OrderTotal-RealTax
        find OrderNo=val(vorderref)
    else
        message "Can't find the order you're using for this transfer"
    endif
    
endif
    

___ ENDPROCEDURE addpay ________________________________________________________

___ PROCEDURE Check Totals _____________________________________________________
getscrap "What's the starting OrderNo?"
select OrderNo ≥ val(clipboard())
getscrap "What's the ending OrderNo?"
selectwithin OrderNo ≤ val(clipboard())
selectwithin CreditCard=""
;message "order number: "+str(OrderNo)+"  "+"1st Payment: "+str(«1stPayment»)
field «1stPayment»
total
___ ENDPROCEDURE Check Totals __________________________________________________

___ PROCEDURE (CommonFunctions) ________________________________________________

___ ENDPROCEDURE (CommonFunctions) _____________________________________________

___ PROCEDURE ExportMacros _____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros are saved to your clipboard!"
___ ENDPROCEDURE ExportMacros __________________________________________________

___ PROCEDURE ImportMacros _____________________________________________________
local Dictionary1,Dictionary2, ProcedureList
Dictionary1=""
Dictionary1=clipboard()
yesno "Press yes to import all macros from clipboard"
if clipboard()="No"
stop
endif
//step one
importdictprocedures Dictionary1, Dictionary2
//changes the easy to read macros into a panorama readable file

 
//step 2
//this lets you load your changes back in from an editor and put them in
//copy your changed full procedure list back to your clipboard
//now comment out from step one to step 2
//run the procedure one step at a time to load the new list on your clipboard back in
//Dictionary2=clipboard()
loadallprocedures Dictionary2,ProcedureList
message ProcedureList //messages which procedures got changed

___ ENDPROCEDURE ImportMacros __________________________________________________

___ PROCEDURE Symbol Reference _________________________________________________
bigmessage "Option+7= ¶  [in some functions use chr(13)
Option+= ≠ [not equal to]
Option+\= « || Option+Shift+\= » [chevron]
Option+L= ¬ [tab]
Option+Z= Ω [lineitem or Omega]
Option+V= √ [checkmark]
Option+M= µ [nano]
Option+<or>= ≤or≥ [than or equal to]"


___ ENDPROCEDURE Symbol Reference ______________________________________________

___ PROCEDURE GetDBInfo ________________________________________________________
local DBChoice, vAnswer1, vClipHold

Message "This Procedure will give you the names of Fields, procedures, etc in the Database"
//The spaces are to make it look nicer on the text box
DBChoice="fields
forms
procedures
permanent
folder
level
autosave
fileglobals
filevariables
fieldtypes
records
selected
changes"
superchoicedialog DBChoice,vAnswer1,“caption="What Info Would You Like?"
captionheight=1”


vClipHold=dbinfo(vAnswer1,"")
bigmessage "Your clipboard now has the name(s) of "+str(vAnswer1)+"(s)"+¶+
"Preview: "+¶+str(vClipHold)
Clipboard()=vClipHold

___ ENDPROCEDURE GetDBInfo _____________________________________________________

___ PROCEDURE .CheckCode1 ______________________________________________________

///*********This is the FileChecker macro in GetMacros
local fileNeeded,folderArray,smallFolderArray,sizeCheck, procList, mostRecentProc

//replace this with whatever file you're error checking
//----------------------//
fileNeeded="members"    //
//----------------------//


case info("files") notcontains fileNeeded and listfiles(folder(""),"????KASX") contains fileNeeded
openfile fileNeeded

case listfiles(folder(""),"????KASX") notcontains fileNeeded
    mostRecentProc=array(info("procedurestack"),1,¬)
    procList=info("procedurestack")
    folderArray=folderpath(folder(""))
    sizeCheck=arraysize(folderArray,":")
    smallFolderArray=arrayrange(folderArray,4,sizeCheck,":")
    message mostRecentProc

//See an example below this codebase of what this looks like 
displaydata "Hi! This is an error checker by Lunar!"
+¶+¶+¶+
"ERROR:"
+¶+
"You are missing the '"+fileNeeded+"' Panorama file in this folder 
and can't continue the '"+mostRecentProc+"'"+¶+" procedure without it. Please move a copy of
'"+fileNeeded+"' to the appropriate folder and try the procedure again"
+¶+¶+¶+
"Pressing 'Ok' will open the Finder to your current folder"
+¶+¶+
"Press 'Stop' will stop this procedure"
+¶+¶+
"You can hit the 'copy' button, and send this to tech-support@fedcoseeds.com if you need help"
+¶+¶+¶+¶+¶+¶+
"THE FOLLOWING LINES ARE TO HELP WITH ERROR CHECKING AND CAN BE DISREGARDED"
+¶+¶+¶+
"folder you're currently running from is: "
+¶+
smallFolderArray
+¶+¶+¶+
"current Pan files in that folder are: "
+¶+
listfiles(folder(""),"????KASX")
+¶+¶+¶+
"currently open files are: "
+¶+
info("files")
+¶+¶+¶+
"last procedures run were"
+¶+
info("procedurestack")
, {title="Missing File!!!!" captionwidth=900 size=17 height=900 width=800}
revealinfinder folder(""),""
stop

defaultcase
window fileNeeded

endcase

/*
Example:

You are missing the 'members' Panorama file in this folder 
and can't continue this procedure without it. Please move a copy of
'members' to the appropriate folder and try the procedure again


folder you're currently running from is: 
Desktop:Panorama:FY45 Panorama Projects:GetMacros:


current Pan files in that folder are: 
GetMacros
GetMacrosDL
GetMacros44


Pressing 'Ok' will open the Finder to your current folder

Press 'Stop' will stop this procedure
*/

debug

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
___ ENDPROCEDURE .CheckCode1 ___________________________________________________

___ PROCEDURE NewSearch/ß ______________________________________________________
/*
This Function is supposed to either 
1. find the customer already in the list
    1a. Call Enter

or

2. add a new line and fill it with this new data

*/

///gets paper_order=0|1 or internet_order=0|1
call ".IsInternetOrder"

if paper_order=-1
message "Sorry, this function only works for internet orders right now. Try CMD-1 or CMD-E instead"
stop
endif 

global vChoice,vFirstInitial,vLastName,vExtracted,vExtracted2,vFirstIntLastName,vCustNum,WinNumber,
vEmail,vPhoneNum, chooseCustomerArray,chooseCustChoice,rayj,MLWildCard,MLWildCard2,orderWildCard,amperArray,thisCustomerLine,
string_zip, conArray

//gets first name and last name
//kept for possible dependencies, 
//but made conArray to replace it
//so I know when other things are calling for the Con
rayj=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1] 
conArray=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1]

waswindow=info("windowname")

window thisFYear+"orders"

if info("formname")="treesinput" and (OrderNo<410000 or OrderNo>420000)
    call ".pool"
    endif

/*This fills the variables for the search*/
EntryDate=today()
ono=«OrderNo»
vzip=Zip
//____added to fix inital zeros disappearing_____///
string_zip=pattern(Zip,"#####")
vd=«C#»  // #Davids old code, leaving in case there are dependencies
vCustNum=«C#»
vEmail=Email
vPhoneNum=Telephone
vChoice=0
place=MAd  //was originally just the number to the end of the address, but PO boxes broke it
vCustNum=«C#»
vEmail=Email
vPhoneNum=Telephone
vChoice=0
MLWildCard=""
MLWildCard2=""
orderWildCard=""
thisCustomerLine=""


//builds a WildCard of the name (see MATCH in the Pan Ref)
    vExtracted=""
    vExtracted=extract(conArray," ",1)
    vFirstInitial=?(vExtracted[1,2] notcontains "", vExtracted[1,2],vExtracted[1,1])
    vExtracted2=""
    vLastName=extract(conArray," ",2)
    orderWildCard=str(vFirstInitial+"*"+vLastName)



field «1stPayment»

///___make sure mailing list is open_____
WinNumber=arraysearch(info("windows"), thisFYear+" mailing list", 1,¶)
if WinNumber=0
    openfile thisFYear+" mailing list"
endif



window thisFYear+" mailing list"

///____can we find them just with the C# they have on the order?
selectall
if vd>0
    find «C#»=vd
    if info("found")=0
        thisCustomerLine=str(«C#»)+" "+Con+" "+Group+¶+MAd+" "+pattern(Zip,"#####")
        YesNo "C# Found. Enter This one?"+¶+thisCustomerLine
        if clipboard() contains "Yes"
            call "enter/e"
            stop
            endif
    endif
endif


Choose:

global ChoiceList, Choice, Options
Choice=""

ChoiceList="First Initial & Last Name
By Email
Phone Number
Mailing Address
Open CMD-F and Let me Search (Pan5 version)
Open CMD-F and Let me Search (Pan6 version)
Add New Customer"

superchoicedialog ChoiceList, Choice, {
caption="No Cust Number Match Found. 
How Would You Like to Search? 
Cancel = Stop ." title="NewSearch/ß (or CMD+Opt+S)" captionstyle=bold captionheight="3" height="250" 
}
if info("dialogtrigger") contains "cancel"
Stop
endif
;message Choice

vChoice=arraysearch(ChoiceList, Choice, 1, ¶)

case vChoice=1
    window thisFYear+" mailing list"
    //try to find the first initial and last name
    //example J*Forester would be the wildcard, and would match names like Jane Forester
    select Con match orderWildCard
        //if it doesn't find it, check the & names
        if info("empty")
            select Con contains "&"
                selectwithin strip(array(Con,1,"&")) match orderWildCard // name left of & matches (Jane Forester & Kathrine Kidd)
                    or strip(array(Con,2,"&")) match orderWildCard // name right of & matches (Katherine Kidd & Jane Forester)
                    or strip(array(Con,1,"&"))+" "+extract(strip(array(Con,2,"&"))," ",2)["- ",-1] match orderWildCard 
                    //name to the left plus the last name to the right matches (Jane & Katherine Forester)
                        if info("empty")
                            message "Unable to find anyone with that name. Please, choose another option"
                                goto Choose
                        endif
        else
            field Con
            sortup
        endif

case vChoice=2
    if vEmail=""
        message "Customer doesn't have an email on order. Reverting to Search."
        goto Choose
    endif
    window thisFYear+" mailing list"
    select email=vEmail
    if info("empty")
        message "No email match found."
        goto Choose
    endif
case vChoice=3
     if vPhoneNum=""
        message "Customer doesn't have an Phone Num on order. Reverting to Search."
        goto Choose
    endif
    window thisFYear+" mailing list"
    select phone=vPhoneNum
        if info("empty")
            message "No phone match found."
            goto Choose
        endif
case vChoice=4
    window thisFYear+" mailing list"
    //confirm this works for all zips
    select MAd contains place and str(Zip) contains str(vzip)
        if info("empty")
        ///_____For PO box entries entered various ways_____///
            //wildcards: Has P+anything or nothing+O+....etc
            //does miss any PO boxes with a " " in front of them
            //so making sure peope can't enter blank as the first character
            //would be wise
            case place match "P"+"*"+"O"+"*"+"box"+"*"
                select MAd match "P"+"*"+"O"+"*"+"box"+"*"
                    selectwithin striptonum(MAd) contains striptonum(place)
                        if info("empty")
                            message "Couldn't autofind this PO Box"
                            goto Choose
                        endif
            defaultcase
                message "No address match found."
                goto Choose
            endcase
        endif
case vChoice=5
    window thisFYear+" mailing list"
    message "Use CMD+E when you select the right customer or CMD+Opt+S to get back here."
    findselect
    if info("dialogtrigger") contains "cancel"
    goto Choose
    else 
        stop
    endif

case vChoice=6
    window thisFYear+" mailing list"
    message "Use CMD+E when you select the right customer or CMD+Opt+S to get back here."
    findselectdialog
    if info("dialogtrigger") contains "cancel"
    goto Choose
    else 
        stop
    endif

case vChoice=7
    window thisFYear+" mailing list"
        insertrecord
        Con=grabdata(thisFYear+"orders",Con) 
        Group=grabdata(thisFYear+"orders",Group) 
        MAd=grabdata(thisFYear+"orders",MAd) 
        City=grabdata(thisFYear+"orders",City) 
        St=grabdata(thisFYear+"orders",St)
        Zip=grabdata(thisFYear+"orders",Zip) 
        adc=lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        SAd=grabdata(thisFYear+"orders",SAd) 
        Cit=grabdata(thisFYear+"orders",Cit) 
        Sta=grabdata(thisFYear+"orders",Sta) 
        Z=grabdata(thisFYear+"orders",Z) 
        phone=grabdata(thisFYear+"orders",Telephone) 
        email=grabdata(thisFYear+"orders", Email)
        adc=lookup("fcmadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        call "numberNeeded"
endcase


window thisFYear+" mailing list"

arrayselectedbuild chooseCustomerArray,¶,"",str(«C#»)+" "+?(Con≠"",Con+", ","")+?(Group≠"",Group+", ","")+MAd+" "+City+", "+St+" "+pattern(Zip,"#####")
        superchoicedialog chooseCustomerArray,chooseCustChoice, {
        captionstyle=bold
        caption="Please choose the appropriate customer or click other search to try something else.
        You can click cancel to look through the selection at any time"
        buttons=OK;OtherSearch:100;Cancel
        captionheight="2" 
        height="600" width="800"
        }
            if info("dialogtrigger") contains "cancel"
                stop
                    endif
            if info("dialogtrigger") contains "other"
                goto Choose
                    endif
            if info("dialogtrigger") contains "ok"
                find exportline() contains chooseCustChoice
                if info("found")=-1
                    thisCustomerLine=str(«C#»)+" "+Con+" "+Group+¶+MAd+" "+string_zip
                    YesNo "Is this is the customer you want?"+¶+thisCustomerLine
                    if clipboard() contains "Yes"
                        call "enter/e"
                        endif
                    endif
                endif
///when does this update the mailing list addresss?
Message "Customer has been entered!"



___ ENDPROCEDURE NewSearch/ß ___________________________________________________

___ PROCEDURE .scrap ___________________________________________________________

___ ENDPROCEDURE .scrap ________________________________________________________

___ PROCEDURE .test.newzipV2 ___________________________________________________
///__________________.custnumber
global addressArray
//this runs when C# is changed
//users are currently using two enters to get this to run during order entry
waswindow=info("windowname")
Global Num
Num=«C#»
ono=OrderNo
addressArray=""
rayj=""


if MAd≠""
    addressArray=MAd+"."+pattern(Zip,"#####")
endif

if Con≠""
    rayj=Con[1," "][1,-2]+" "+Con["- ",-1][2,-1]
endif

if «C#»=0
    window newyear+" mailing list"
    call ".newzip"
    stop
endIf 

window newyear+" mailing list"
find «C#»=Num

if info("found")=0
    call "getzip/Ω"
else
    window waswindow
    call ".customerfill"   
endif


////__________________.NewZIp
fileglobal listzip, thiszip, findher, findzip, findcity, newcity, findname,findname1, findname2, thisname, firstname, lastname
serverlookup "off" 
;waswindow=info("windowname")
listzip=""
thiszip=""
newcity=""
again:
findher=addressArray


supergettext findher, {caption="Enter Address.Zip" height=100 width=400 captionfont=Times captionsize=14 captioncolor="cornflowerblue"
    buttons="Find;Redo;Cancel"}
    if info("dialogtrigger") contains "Find"
        findzip=extract(findher,".",2)
        findzip=strip(findzip)
            if length(findzip)=4
            findzip="0"+findzip
            endif
        findcity=extract(findher,".",1)
        liveclairvoyance findzip,listzip,¶,"",thisFYear+" mailing list",pattern(Zip,"#####"),"=",str(«C#»)+¬+rep(" ",7-length(str(«C#»)))+Con+rep(" ",max(20-length(Con),1))+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####"),0,0,""
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
        liveclairvoyance lastname,findname1,¶,"",thisFYear+" mailing list",Con,"contains",Con+¬+MAd+¬+City+¬+St+¬+pattern(Zip,"#####")+¬+phone,0,0,""
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
            window thisFYear+" mailing list"
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


___ ENDPROCEDURE .test.newzipV2 ________________________________________________

___ PROCEDURE break ____________________________________________________________


        Pool = 20

        Pool = 22

        call ".pool"


    call ".memberdisc"
    MemDisc = 0
    call ".retotal"


Field SAd 

    Field Telephone
 
        field Email
    
                Field «1stPayment»
      
                Field Sub1
                editcell
                Field Sub2
                editcell
                field «1stPayment»
 
      
        EditCellStop
       
            Field «1stPayment»
___ ENDPROCEDURE break _________________________________________________________

___ PROCEDURE TestCode _________________________________________________________
displaydata pattern(Zip,"#####")
___ ENDPROCEDURE TestCode ______________________________________________________

___ PROCEDURE GetWindowSize ____________________________________________________
global newrec, rectangle1,RecTop,RecLeft,RecHeight,RecWidth,whichWin,winList2

winList2=info("windows")
superchoicedialog winList2,whichWin,“caption="Which Window do you want the size of?"”
window (whichWin)
rectangle1=info("windowrectangle")
RecTop=rtop(rectangle1)
RecLeft=rleft(rectangle1)
RecHeight=rheight(rectangle1)
RecWidth=rwidth(rectangle1)

newrec=str(RecTop)+","+str(RecLeft)+","+str(RecHeight)+","+str(RecWidth)
message "You now have the Top, Left, Height, and Width of the window. You can use the setwindow command with these numbers"
clipboard()=newrec
//top,left,height,width

___ ENDPROCEDURE GetWindowSize _________________________________________________

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

___ PROCEDURE .getequations ____________________________________________________
global eqArray
arraybuild eqArray, ¶,"", ?(Equation≠"","~~"+«Field Name»+"~~"+¶+Equation+¶,"")
clipboard()=eqArray
___ ENDPROCEDURE .getequations _________________________________________________

___ PROCEDURE .Test ____________________________________________________________
waswindow=info("windowname")
//don't do any of the following on addresschecker
    if info("formname") contains "addresschecker"
    stop
    endif
//don't do anything but move fields if I used Tab to get here
    If info("trigger")="Key.Tab"
    Field SAd
    EndIf

//GetNumber
if «C#»=0
    message "This Customer Must have a Customer # to use this address updater"
    stop
endif
ID=«C#»

//________Update Files with New Address from Order______//
If info("trigger")="Key.Return"
YesNo "Update "+thisFYear+" "+"mailing list & customer_history Mailing Address?"
        If clipboard()="Yes"
      window thisFYear+" mailing list"
            If «C#»≠ID
            if info("selected")<info("records")
            selectall
            endif
            Find «C#»=ID
            EndIf
        MAd=grabdata(thisFYear+"orders", MAd)
        City=grabdata(thisFYear+"orders", City)
        St=grabdata(thisFYear+"orders", St)
        Zip=grabdata(thisFYear+"orders", Zip)
        adc=lookup("newadc","Zip3",pattern(Zip,"#####")[1,3],"adc",0,0)
        Field «C#»
        window "customer_history:secret"
        Find «C#»=ID
        MAd=grabdata(thisFYear+"orders", MAd)
        City=grabdata(thisFYear+"orders", City)
        St=grabdata(thisFYear+"orders", St)
        Zip=grabdata(thisFYear+"orders", Zip)
        window waswindow
        EndIf
        if info("formname")=thisFYear+"orders:addresschecker"
stop
endif
Field SAd
EndIf

___ ENDPROCEDURE .Test _________________________________________________________

___ PROCEDURE .IsInternetOrder _________________________________________________
///__tests the first two digits to see if it's an internet order____///
global paper_orders_array,paper_order,internet_order
paper_orders_array="30,31,40,41,50,51,60,61,70"
internet_order=0
paper_order=0

if  paper_orders_array notcontains str(OrderNo)[1,2]
        internet_order=-1
    else
        paper_order=-1
endif
___ ENDPROCEDURE .IsInternetOrder ______________________________________________

___ PROCEDURE .FromBranch ______________________________________________________
//this came from next/1 and is simplified just to check what branch something is from for other things
//there's a text version and a num version

global intOrder1, fromBranch, BranchNum

intOrder1=int(OrderNo)
fromBranch=""
BranchNum=0

case intOrder1 ≥ 700000
    fromBranch="Seeds"
    BranchNum=1
case intOrder1 ≥ 600000 and intOrder1 < 700000
    fromBranch="OGS;POE"
    BranchNum=2
case intOrder1 ≥ 500000 and intOrder1 < 600000
    fromBranch="OGS;Bulbs"
    BranchNum=3
case intOrder1 ≥ 400000 and intOrder1 < 500000
    fromBranch="Trees"
    BranchNum=4
case intOrder1 ≥ 300000 and intOrder1 <400000
    fromBranch="OGS"
    BranchNum=5
endcase

___ ENDPROCEDURE .FromBranch ___________________________________________________
