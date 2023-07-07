

Global New, Num, Numb, enter, place, ID, CC, ED, credit, orders, ship, Flag, vFacWindow, ono, folder_files, waswindow

waswindow = info("windowname")

folder_files = listfiles(folder(""),"????KASX")


Num=0
zoomWindow 23,0,250,1496,""
vFacWindow = info("windowname")
NoShow
openFile "46 mailing list"
    SelectAll
    
    call "listsortcomplete/1"
    save
    //CloseFile
    
window vFacWindow

OpenFile "&&46 mailing list"
OpenForm "view orders blue box"
   zoomwindow 285,0,494,802,""
   OpenFile "&&46 mailing list"

save
EndNoShow

OpenSecret "newadc"
    
;Openfile "customer_history"
;    zoomwindow 783,0,410,647,""
;   save
;    CloseFile

Openfile "46orders"
    goform "facilitator view"
    zoomwindow 25,1159,764,760,""
    save



   ; OpenForm "functions"
  ; zoomwindow 885,850,153,241,""
  
Openfile "46shipping"
    goform "packages"
    zoomwindow 868,1223,326,701,""
    save

if folder_files contains "46seedstally-winter"    
Openfile "46seedstally-winter"
    zoomwindow 83,959,951,932,""
    field OrderNo
    sortup
    save
    endif
    
    if folder_files contains "45seedstally-winter"    
Openfile "45seedstally-winter"
    zoomwindow 83,959,951,932,""
    field OrderNo
    sortup
    save
    endif

if folder_files contains "46seedstally"
openfile "46seedstally"
    zoomwindow 83,959,951,932,""
    field OrderNo
    sortup
    save
    endif

Openfile "46ogstally"

    // window "46ogstally:bulbspagecheck"
   // zoomwindow 163,903,899,870,""
    field OrderNo
    sortup
    save

if folder_files contains "46treestally"
Openfile "46treestally"
   zoomwindow 118,931,822,862,""
    field OrderNo
    sortup
    save
    endif
    
    if folder_files contains "45treestally"
Openfile "45treestally"
   zoomwindow 118,931,822,862,""
    field OrderNo
    sortup
    save
    endif
    
        if folder_files contains "46ogstally"
Openfile "46ogstally"
   zoomwindow 118,931,822,862,""
    field OrderNo
    sortup
    save
    endif

    if folder_files contains "45ogstally"
Openfile "45ogstally"
   zoomwindow 118,931,822,862,""
    field OrderNo
    sortup
    save
    endif




;Openfile "46bulbstally"
;   zoomwindow 198,961,822,862,""
;    field OrderNo
;    sortup
;    save
openfile "ycitylist"
;window  "46 facilitator mailing list:functions"
window  "46 facilitator mailing list:view orders blue box"
message "Faciiltation Mailing List has finished opening."