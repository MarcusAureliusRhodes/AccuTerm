Sub Main

! Copyright 1994-2020 Marcus Rhodes <marcus1@thinqware.com>

!           This program is free software; You can redistribute it and/or
!           modify it under the terms of the GNU general  public  license
!           version 3 as published by the Free Software Foundation.

! Modified: 12/01/2018 17:04:52 by Marcus

! Platform: Any Pick; Any OS; AccuTerm; Any emulation

! Function: Provide a sample VBA macro for EMV_ATW_VBA_RUN

! Syntax  : Enter "EMV_ATW_VBA_RUN" at TCL, and follow the prompts.

! Method  : This is a prototype VBA script, as well as a  sample  of  the
!           sort of VBA code you're most likely to  need  for  automating
!           the operation of a site's application.  And that may  be  for
!           automated testing  purposes,  or  for  creating  a  robot  to
!           simplify repetitive tasks without having to  reverse-engineer
!           a lot of Pick Basic code, and build a new program to automate
!           those tasks.

!           This script will obviously not run as-is.  It must  first  be
!           pre-processed, and then sent to  AccuTerm  for  execution  by
!           EMV_ATW_VBA_RUN.  (See EMV_ATW_VBA_RUN for more.)

!           Whatever application you adapt this script to, the first step
!           should be  to  copy  it  to  a  new  script.  Keep  this  one
!           inviolate.  Then populate the Prompts array (below)  of  your
!           new script with all the prompts a user might encounter during
!           run-time, longest first, probably, except for the  first  few
!           debugger/system prompt strings.

!           The first few Prompts (exactly how many is  up  to  you) (but
!           they should at least encompass any that could appear on  your
!           platform) should probably never change or move simply because
!           they represent catastrophic failures, which, like it or  not,
!           could happen, and, if present, could  even  appear  on-screen
!           along with one or more of the other 'valid' prompts, so these
!           failure conditions must always be checked for first, in order
!           to prevent false-positives of the other prompts.

!           Moreover, this/your test-script ought to be able to deal with
!           such failures in an orderly fashion.  And this script already
!           contains sample logic for handling such  situations,  and  it
!           would probably be a trivial matter to  adapt  that  logic  to
!           your own script.

! Upcoming: Change is the only constant. -- Heraclitus

   Dim Prompts$( 110 )

   Prompts(   0 ) = "jBASE debugger->"
   Prompts(   1 ) = "sh CMI ~ -->"
   Prompts(   2 ) = "Enter I to Ignore, R to Retry , Q to Quit :"
   Prompts(   3 ) = "<E>nter New Email Address  <Enter>=Accept Email Address Displayed  X=Go Back"
   Prompts(   4 ) = "<ENTER> = page, B-ack, X-exit, C-hange, P-ackMsg, RU-Rules, S-hipIns, F-ile,"
   Prompts(   5 ) = "<ENTER> = next page  A = Attn / RU = Rules / H = help / B = Back / Q = quit"
   Prompts(   6 ) = "<ENTER> = page, B-ack, X = exit, C-hange, RU = Rules, P-ackMsg, F-ile"
   Prompts(   7 ) = "ENTER Customer Number  H=Help  S=Search  A=Add New Customer  Q=Quit"
   Prompts(   8 ) = "Sales Order Confirmation has been emailed...press enter to continue"
   Prompts(   9 ) = "<ENTER> = next page  A = Attn / RU = Rules / H = help / B = Back"
   Prompts(  10 ) = "<ENTER> continue, # Add or Change, H-elp, Q-uit, X=Go Back:"
   Prompts(  11 ) = "<ENTER> to continue / <#> to Change / <H>elp / <B>=Back :"
   Prompts(  12 ) = "H=help / L=list / X=back / Q=quit order / END=finished"
   Prompts(  13 ) = "<ENTER> continue, # Add or Change, H-elp, Q-uit, X=Go"
   Prompts(  14 ) = "H=help / L=list / Q=quit order/ X=back / END=finished"
   Prompts(  15 ) = "Do you want to <F>ax or <E>mail order confirmation ?"
   Prompts(  16 ) = "ENTER Customer Number  H=Help  S=Search  Q=Quit"
   Prompts(  17 ) = "OK to File and Send Email Confirmation (Y/N)?"
   Prompts(  18 ) = "ENTER Customer Number  H=Help  S=Search"
   Prompts(  19 ) = "The product is on temporary hold until"
   Prompts(  20 ) = "/ RU=Rules / H=help / B=Back / Q=quit"
   Prompts(  21 ) = "Are you SURE you want to Quit  Y/N :"
   Prompts(  22 ) = "RU=Rules / H=help / B=Back / Q=quit"
   Prompts(  23 ) = "Not a valid C.M.I. Customer Number"
   Prompts(  24 ) = "*** Enter Y or N to continue **"
   Prompts(  25 ) = "*** Press ENTER to continue ***"
   Prompts(  26 ) = "H=help / S=PROD Search / X=back"
   Prompts(  27 ) = "H=help <Enter>=NA X=back Q=Quit"
   Prompts(  28 ) = "Send Invoice to Email Address:"
   Prompts(  29 ) = "Send Copy of Invoice? (Y/N):"
   Prompts(  30 ) = "Place Order on Hold Y/<N>? :"
   Prompts(  31 ) = "You may not select customer"
   Prompts(  32 ) = "The UOM on this product is"
   Prompts(  33 ) = "Custom Device tracking #:"
   Prompts(  34 ) = "Take another Order  Y/N ?"
   Prompts(  35 ) = "Current or Default (C/D)"
   Prompts(  36 ) = "Press ENTER to continue"
   Prompts(  37 ) = "to purchase product in"
   Prompts(  38 ) = "is on the PRICE PAUSE"
   Prompts(  39 ) = "Logging you off....."
   Prompts(  40 ) = "CANNOT USE OBSOLETE"
   Prompts(  41 ) = "is transitioning to"
   Prompts(  42 ) = "Accept Hold? (Y/N)"
   Prompts(  43 ) = "Continue? (Y/N)"
   Prompts(  44 ) = "Expiration Date"
   Prompts(  45 ) = "Are you sure ?"
   Prompts(  46 ) = "records copied"
   Prompts(  47 ) = "Base Reason :"
   Prompts(  48 ) = "X=back Q=Quit"
   Prompts(  49 ) = "OK to SAVE? :"
   Prompts(  50 ) = "Enter Option"
   Prompts(  51 ) = "END=finished"
   Prompts(  52 ) = "EMAIL NOTES:"
   Prompts(  53 ) = "Base Reason"
   Prompts(  54 ) = "Q = quit"
   Prompts(  55 ) = "NA = NA"
   Prompts(  56 ) = "Option:"
   Prompts(  57 ) = "Q=Quit"
   Prompts(  58 ) = "X=back"
   Prompts(  59 ) = "2.  :"
   Prompts(  60 ) = "1.  :"
   Prompts(  61 ) = "TO :"
   Prompts(  62 ) = ""

   ! The final element of all the possible prompts must always be a  null
   ! string.  Immediately following that come whatever strings  are  best
   ! made globally available via this array.

   ! I'll be skipping several places just to minimize renumbering  should
   ! any additions be required in the future.

   Prompts(  70 ) = "--------------------------------------------------------------------------------"
   Prompts(  71 ) = "================================================================================"
   Prompts(  72 ) = "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"

   ! Elements beyond this are open for  other  'global'  variables  to  be
   ! 'equated' to.

   ! I'll be skipping several here, too.

   EQU Action   TO Prompts(  80 )
   EQU WaitTime TO Prompts(  81 )
   EQU SnapShot TO Prompts(  82 )
   EQU TestName TO Prompts(  83 )

   EQU CaptPath TO Prompts(  84 )
   EQU ConfPath TO Prompts(  85 )
   EQU CustPath TO Prompts(  86 )
   EQU ProdPath TO Prompts(  87 )
   EQU SummPath TO Prompts(  88 )

   EQU CustCode TO Prompts(  89 )
   EQU ProdCode TO Prompts(  90 )

   EQU DomVsInt TO Prompts(  91 )
   EQU OrderHdr TO Prompts(  92 )

   EQU Beg_Time TO Prompts(  94 )
   EQU Cust_Max TO Prompts(  95 )
   EQU Prod_Max TO Prompts(  96 )

   EQU GameOver TO Prompts(  97 )

   EQU Reaction TO Prompts(  98 )
   EQU SummText TO Prompts(  99 )
   EQU End_Time TO Prompts( 100 )
   EQU ElpsTime TO Prompts( 101 )
   EQU PgBrLine TO Prompts( 102 )
   EQU UserName TO Prompts( 103 )
   EQU PurchNum TO Prompts( 104 )
   EQU PurchInt TO Prompts( 105 )
   EQU OrderNum TO Prompts( 106 )

   EQU PageLine TO Prompts( 107 )
   EQU WarnLine TO Prompts( 108 )
   EQU StepLine TO Prompts( 109 )

   ! These next 2 equates provide a sort of Babelizer, but only to  shrink
   ! the size of the script passed to AccuTerm rather than  to  complicate
   ! reverse-engineering.  In fact, the entire point of this  exercize  is
   ! readability/maintainability.

   EQU Prompts  TO p
   EQU Prompts$ TO p$

   ! And these eliminate the need to make the file-handles global.

   EQU CaptFile TO 1
   EQU ConfFile TO 2
   EQU CustFile TO 3
   EQU ProdFile TO 4
   EQU SummFile TO 5

   Dim w As Session

   InitSession.Activate

   Set w = ActiveSession

   w.Reset atResetTerminal

!  dim w as accutermclasses.session

!  set w = activesession

!  accuterm.activate

   w.InputMode = 0

!  HomePath = GetFilePath$( "*", "*", MacroDir$, "Locate a directory you have write-access to.", 0 )

   ! HomePath = GetFilePath$( [DefName$], [DefExt$], [DefDir$], [Title$], [Option])

   ! DefName$  Set the initial File Name in the to this string value. If this is omitted then *.DefExt$ is used.
   ! DefExt$   Initially show files whose extension matches this string value. (Multiple extensions can be specified by using ";" as the separator.) If this is omitted then * is used.
   ! DefDir$   This string value is the initial directory. If this is omitted then the current directory is used.
   ! Title$    This string value is the title of the dialog. If this is omitted then ''Open" is used.
   ! Option
   !      0    Only allow the user to select a file that exists.  (Default)
   !      1    Confirm creation when the user selects a file that does not exist.
   !      2    Allow the user to select any file whether it exists or not.
   !      3    Confirm overwrite when the user selects a file that exists.
   !     +4    Selecting a different directory changes the application's current directory.

!  HomePath = OpenFilename$( "Locate a directory you have write-access to.", "*.*" )
!  HomePath = SaveFilename$( "Locate a directory you have write-access to.", "*.*" )

   HomePath = cstr( Environ( "USERPROFILE" ) )
   CaptPath = HomePath & "\Documents"
   ConfPath = CaptPath & "\UserInfo.txt"
   CustPath = CaptPath & "\CustList.txt"
   ProdPath = CaptPath & "\ProdList.txt"
   SummPath = CaptPath & "\SummTemp.txt"

   Close ;! any/all open files

   GameOver = "False"
   WaitTime = "9"

   Begin Dialog UserDialog 280, 230, "Test Menu"

      PushButton 10,  20, 260,  22, "Reset databases"
      PushButton 10,  50, 260,  22, "Xfer approvals from GARS to GPW"
      PushButton 10,  80, 260,  22, "Run Test 005 (Domestic Order)"
      PushButton 10, 110, 260,  22, "Run Test 006 (Foreign Order)"
      PushButton 10, 140, 260,  22, "Run Test 012 (Nightly Purge)"
      PushButton 10, 170, 260,  22, "Help"
      PushButton 10, 200, 260,  22, "Quit"

   End Dialog

   Dim MainMenu As UserDialog

   MenuDone = False

   Call OpenCapt( w, Prompts )
   Call NextStep( w, "MENU", Prompts )

   Do

      UserSays = CStr( Dialog( MainMenu, 7 ) )

      Select Case UserSays

         Case 1

            Call ResetData( w, Prompts )

         Case 2

            Call XferData( w, Prompts )

         Case 3

            DomVsInt = "1"
            OrderHdr = "Domestic Order"
            TestName = "005"

            Call DomesticOrder( w, Prompts )

         Case 4

            DomVsInt = "11"
            OrderHdr = "Foreign Order"
            TestName = "006"

            Call DomesticOrder( w, Prompts )

         Case 5

            DomVsInt = "1"
            OrderHdr = "Nightly Purge"
            TestName = "012"

            Call DomesticOrder( w, Prompts )

         Case 6

            Call ShowHelp

         Case Else

            MenuDone = True

      End Select

   Loop Until MenuDone

End Sub

Sub StartVariables( w, Prompts )
End Sub

Sub ResetData( w, Prompts )

   TestName = "Reset_Data"

   Call OpenCapt( w, Prompts )

   Call NextStep( w, "TCL"                    , Prompts )
   Call NextStep( w, "COPY CMI-WACIM-BAK * (O", Prompts )

   WaitTime = "300"

   Call NextStep( w, "(CMI-WACIM", Prompts )

   WaitTime = "9"

   Call NextStep( w, "cp -f /data/files/cmi/GCR.DATA.BAK /data/files/cmi/GCR.DATA"                , Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/GCR.DATA.BAK /data/files/cmi/GCR.DATA"                , Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/GCR.DEFINITION.BAK /data/files/cmi/GCR.DEFINITION"    , Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/WAC-CLM.BAK /data/files/cmi/WAC-CLM"                  , Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/WAC-IM-AVAIL.BAK /data/files/cmi/WAC-IM-AVAIL"        , Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/WAC-RESTRICTIONS.BAK /data/files/cmi/WAC-RESTRICTIONS", Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/WACCM.BAK /data/files/cmi/WACCM"                      , Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/WACCM-SHIP-ADDRS.BAK /data/files/cmi/WACCM-SHIP-ADDRS", Prompts )
   Call NextStep( w, "cp -f /data/files/cmi/WACIR.BAK /data/files/cmi/WACIR"                      , Prompts )

   WaitTime = "60"

   Call NextStep( w, "cp -f /data/files/cmi/WACO.BAK /data/files/cmi/WACO", Prompts )

   WaitTime = "99"

   Call NextStep( w, "MENU", Prompts )

   Close #CaptFile

   Shell( "C:\Windows\notepad.exe " & CaptPath, vbMaximizedFocus )

End Sub

Sub XferData( w, Prompts )
End Sub

Sub DomesticOrder( w, Prompts )

   Beg_Time = CStr( Time )
   Cust_Max = CStr( -1 )
   Prod_Max = CStr( -1 )

   If FileExists( CustPath ) Then

      Open CustPath For Input As #CustFile

      Do

         Input #CustFile, CustCode

         Cust_Max = CStr( CInt( Cust_Max ) + 1 )

      Loop Until CustCode = "EOF"

      Close #CustFile

   Else

      MsgBox "Customer File does not exist. Download from https://confluence.cookgroup.nao/display/JTOCI/jBTO+GARS+Automation to " & CustPath, vbOkOnly, "Error"

      GameOver = "True"

   End If

   If FileExists( ProdPath ) Then

      Open ProdPath For Input As #ProdFile

      Do

         Input #ProdFile, ProdCode

         Prod_Max = CStr( CInt( Prod_Max ) + 1 )

      Loop Until ProdCode = "EOF"

      Close #ProdFile

   Else

      MsgBox "Product File does not exist. Download from https://confluence.cookgroup.nao/display/JTOCI/jBTO+GARS+Automation to " & ProdPath, vbOkOnly, "Error"

      GameOver = "True"

   End If

   TestName = "Domestic_Order_(005)"

   Call OpenCapt( w, Prompts )

   Call NextStep( w, DomVsInt, Prompts )
   Call NextStep( w, "1"     , Prompts )

   Open SummPath For Output As #SummFile
   Open CustPath For Input  As #CustFile

   For Cust_Idx = 1 To Cust_Max

      SummText = ""

      Input #CustFile, CustCode

      Call CheckCust( w, Prompts )

      If CInt( Reaction ) = 1 Then

         Call EnterCust( w, Prompts )

         Open ProdPath For Input As #ProdFile

         For Prod_Idx = 1 To Prod_Max

            Input #ProdFile, ProdCode

            Call CheckProd( w, Prompts )

            If CInt( Reaction ) = 1 Then

               Call EnterOrder( w, Prompts )

            End If

            Call UpdateSummary( w, Prompts )

         Next Prod_Idx

         Close #ProdFile

         Call NextStep( w, "Q", Prompts )
         Call NextStep( w, "Y", Prompts )
         Call NextStep( w, "Y", Prompts )

      End If

   Next Cust_Idx

   Close #CustFile

   Call NextStep( w, "Q"  , Prompts )
   Call NextStep( w, "Q"  , Prompts )
   Call NextStep( w, "OFF", Prompts )

   End_Time = CStr( Time )
   ElpsTime = CStr( CInt( End_Time ) - CInt( Beg_Time ) )

   Print #CaptFile, PageLine
   Print #CaptFile, ""
   Print #CaptFile, "                    " & OrderHdr & " (" & TestName & ") Test Summary Report"
   Print #CaptFile, ""
   Print #CaptFile, "Report Path : " & CaptPath
   Print #CaptFile, ""
   Print #CaptFile, "Username    : " & UserName
   Print #CaptFile, "Start Time  : " & Format( CInt( Beg_Time ), "hh:mm:ss" )
   Print #CaptFile, "End Time    : " & Format( CInt( End_Time ), "hh:mm:ss" )
   Print #CaptFile, "Elapsed Time: " & Format( CInt( ElpsTime ), "hh:mm:ss" )
   Print #CaptFile, ""
   Print #CaptFile, Cust_Max & " customers"
   Print #CaptFile, Prod_Max & " products"
   Print #CaptFile, ( CInt( Cust_Max ) * CInt( Prod_Max ) ) & " potential tests"
   Print #CaptFile, ""
   Print #CaptFile, "Actual"
   Print #CaptFile, "Test # Customer   Product        Outcome"
   Print #CaptFile, PageLine

   Print #SummFile, "EOF"
   Close #SummFile

   Open SummPath For Input As #SummFile

   Test_Num = 0

   Do

      Input #SummFile, Temp$

      If Temp$ = "EOF" Then

         Exit Do

      Else

         Test_Num = Test_Num + 1

         Print #CaptFile, Right( "      " & Str( Test_Num ) & "  ", 7 ) & Temp$

      End If

   Loop

   Print #CaptFile, PageLine

   Close #CaptFile
   Close #SummFile

   Shell( "C:\Windows\notepad.exe " & CaptPath, vbMaximizedFocus )

End Sub

Sub ShowHelp

   HelpText$ =             "Reset databases ... copies the `snapshot` database files over the current files in order to restore the system to a standard state for testing." & vbCr & vbCr
   HelpText$ = HelpText$ & "Xfer approvals from GARS to GPW ... transfers pending approvals from GARS to GPW so orders can be created with them." & vbCr & vbCr
   HelpText$ = HelpText$ & "Run Test 005 (Domestic Order) ... starts the domestic product ordering test script." & vbCr & vbCr
   HelpText$ = HelpText$ & "Run Test 006 (Foreign Order) ... starts the international product ordering test script." & vbCr & vbCr
   HelpText$ = HelpText$ & "Run Test 012 (Nightly Purge) ... starts the test of the nightly purge using the domestic order script."

   MsgBox HelpText$, vbOkOnly, "The buttons do the following..."

End Sub

Sub CheckCust( w, Prompts )

   Call NextStep( w, CustCode, Prompts )

   Select Case CInt( Reaction )

      Case 2

         Reaction = "1"

      Case 3

         SummText = SummText & "Obsolete customer; "

         Call NextStep( w, "", Prompts )

      Case 4

         SummText = SummText & "Invalid customer; "

         Call NextStep( w, "", Prompts )

      Case 5

         w.SetSelection 20,9,62,16

         SummText = SummText & Replace( Replace( Replace( w.Selection, vbLf, " " ), vbCr, " " ), "  ", " ", 1, 100 ) & "; "

         Call NextStep( w, "", Prompts )

         Reaction = "2"

      Case 6

         SummText = SummText & "Customer on price-pause; "

         Call NextStep( w, "", Prompts )

      Case 7

         Call NextStep( w, "", Prompts )

         Select Case Reaction

            Case "2"

               Call NextStep( w, "Q", Prompts )

         End Select

   End Select

   ProdCode = ""

   Call UpdateSummary( w, Prompts )

End Sub

Sub EnterCust( w, Prompts )

   Call NextStep( w, ""                            , Prompts )
   Call NextStep( w, "JOHN SMITH"                  , Prompts )
   Call NextStep( w, "RECEIVING"                   , Prompts )
   Call NextStep( w, "3171234567"                  , Prompts )
   Call NextStep( w, ""                            , Prompts )
   Call NextStep( w, "mrhodes@robertetusa.com", Prompts )
   Call NextStep( w, ""                            , Prompts )
   Call NextStep( w, PurchNum                      , Prompts )

   If Reaction = "2" Then

      Call NextStep( w, "Y", Prompts )

   End If

   Call NextStep( w, "V"               , Prompts )
   Call NextStep( w, "1234567890123456", Prompts )
   Call NextStep( w, "JOHN SMITH"      , Prompts )
   Call NextStep( w, "0817"            , Prompts )
   Call NextStep( w, ""                , Prompts )
   Call NextStep( w, ""                , Prompts )

   If Reaction = "2" Then

      Reaction = "1"

   End If

   Call NextStep( w, "5", Prompts )

   If Reaction = "2" Or Reaction = "3" Then

      Reaction = "1"

   End If

End Sub

Sub StoreInfo( w, Prompts )

   PurchInt = CStr( CInt( PurchInt ) + 1 )
   PurchNum = Replace$( UserName & PurchInt, " ", "" )

   Open ConfPath For Output As #ConfFile

   Print #ConfFile, UserName
   Print #ConfFile, PassWord$
   Print #ConfFile, PurchInt
   Print #ConfFile, ATCFPath$

   Close ConfFile

End Sub

Sub CheckProd( w, Prompts )

   Call CheckProdPart2( w, ProdCode, Prompts )

End Sub

Sub CheckProdPart2( w, Dummy$, Prompts )

   ProdCode = Dummy$

   Call NextStep( w, ProdCode, Prompts )

   Select Case Reaction

      Case "2"

         Action = "C"

         Call CheckProdPart2( w, ProdCode, Prompts )

         Reaction = "1"

      Case "3"

         Wait 0.2

         w.SetSelection 0,21,80,22

         Call AppendSelection( w, Prompts )

      Case "4"

         SummText = SummText & "Product on hold; "

         Action = "Y"

         Call CheckProdPart2( w, ProdCode, Prompts )

      Case "5"

         Wait 0.2

         w.SetSelection 20,9,62,16

         Call AppendSelection( w, Prompts )

         Action = ""

         Call CheckProdPart2( w, ProdCode, Prompts )

      Case "6"

         SummText = SummText & "Custom Device Tracking #; "

         Call NextStep( w, "100", Prompts )
         Call NextStep( w, ""   , Prompts )
         Call NextStep( w, "5"  , Prompts )

      Case "7"

         If DomOrInt = "1" Then

            Reaction = "1"

         Else

            Reaction = "1"

         End If

      Case "8"

         Call NextStep( w, "Y", Prompts )

      Case "9"

   End Select

End Sub

Sub EnterOrder( w, Prompts )

   Call NextStep( w, ""   , Prompts )
   Call NextStep( w, "END", Prompts )

   If Reaction = "2" Then

      Call NextStep( w, ""          , Prompts )
      Call NextStep( w, "U"         , Prompts )
      Call NextStep( w, "1"         , Prompts )
      Call NextStep( w, "JOHN SMITH", Prompts )
      Call NextStep( w, "3171234567", Prompts )

   End If

   Call NextStep( w, ""                , Prompts )
   Call NextStep( w, "N"               , Prompts )
   Call NextStep( w, "Robertet Testing", Prompts )
   Call NextStep( w, ""                , Prompts )

   If DomVsInt = "11" Then

      Call NextStep( w, "", Prompts )
      Call NextStep( w, "", Prompts )

   End If

   Call NextStep( w, "" , Prompts )
   Call NextStep( w, "F", Prompts )

   Select Case Reaction

      Case "2"

         Call NextStep( w, "C", Prompts )

      Case "3"

         Call NextStep( w, "Y", Prompts )

   End Select

   Call NextStep( w, "Y"                      , Prompts )
   Call NextStep( w, "mrhodes@robertetusa.com", Prompts )
   Call NextStep( w, "Y"                      , Prompts )

   w.SetSelection 20,16,30,16

   OrderNum = w.Selection

   If Reaction = "2" Then

   End If

   Call NextStep( w, "E"               , Prompts )
   Call NextStep( w, ""                , Prompts )
   Call NextStep( w, "ROBERTET TESTING", Prompts )
   Call NextStep( w, ""                , Prompts )
   Call NextStep( w, "Y"               , Prompts )
   Call NextStep( w, ""                , Prompts )
   Call NextStep( w, "Y"               , Prompts )
   Call NextStep( w, CustCode          , Prompts )

   Call EnterCust( w, Prompts )

   SummText = SummText & "Successfully ordered (" & OrderNum & "); "

End Sub

Sub OpenCapt( w, Prompts )

   CaptPath = CaptPath & Format(   Year( Date ), "0000" )
   CaptPath = CaptPath & Format(  Month( Date ),  "00"  )
   CaptPath = CaptPath & Format(    Day( Date ),  "00"  )
   CaptPath = CaptPath & Format(   Hour( Time ),  "00"  )
   CaptPath = CaptPath & Format( Minute( Time ),  "00"  )
   CaptPath = CaptPath & Format( Second( Time ),  "00"  )
   CaptPath = CaptPath & "_Marcus_Test_000.txt"

   Open CaptPath For Output As #CaptFile

End Sub

Sub AppendSelection( w, Prompts )

   SummText = SummText & Replace( Replace( Replace( w.Selection, vbLf, " " ), vbCr, " " ), "  ", " ", 1, 100 ) & "; "

End Sub

Sub NextStep( w, Dummy, Prompts )

   Action = Dummy

   If GameOver = "False" Then

      w.SetSelection 0,0,80,24

      If Action = "Do Nothing!" Then

         SnapShot = w.Selection

         Print #CaptFile, SnapShot
         Print #CaptFile, PageLine
         Print #CaptFile, ""
         Print #CaptFile, PageLine

      Else

         w.Output Action

         SnapShot = w.Selection

         Print #CaptFile, SnapShot
         Print #CaptFile, PageLine
         Print #CaptFile, Action & "{Enter}"
         Print #CaptFile, PageLine

         w.Output vbCr

      End If

      z = w.WaitFor( 0, CInt( WaitTime ), " ", ":", ">", "[", Chr$( 27 ) )

      Wait 0.2

      SnapShot = w.Selection

      For X = 0 To 99

         FindThis$ = Prompts( X )

         If FindThis$ = "" Then

            Reaction = "-1"

            Exit For

         Else

            If InStr( 1, Replace( Replace( SnapShot, vbCr, "|" ), vbLf, "|" ), FindThis$ ) Then

               Reaction = CStr( X )

               Exit For

            End If

         End If

      Next X

      Select Case Reaction

         Case "-1"

            Call BreakOff( w, Prompts )

         Case "0"

            Call BreakOff( w, Prompts )

      End Select

   End If

End Sub

Sub BreakOff( w, Prompts )

   w.SetSelection 0,0,80,24

   Print #CaptFile, w.Selection
   Print #CaptFile, PageLine
   Print #CaptFile, "!!!!!!!!!TIMED OUT WAITING FOR ONE OF THE FOLLOWING RESPONSES TO APPEAR!!!!!!!!!"

   For X = 0 To 99

      If Prompts( X ) = "" Then

         Exit For

      Else

         Print #CaptFile, "Prompts(" & Right$( "  " & Str$( X ), 2 ) & ") = `" & Prompts( X ) & "`"

      End If

   Next X

   Print #CaptFile, Prompts( 65 )
   Print #CaptFile, PageLine

   SummText = "ABORTED!"

   Call UpdateSummary( w, Prompts )

   Call NextStep( w, ChrW$( 3 ), Prompts )
   Call NextStep( w, "Q"       , Prompts )
   Call NextStep( w, "Y"       , Prompts )
   Call NextStep( w, "Q"       , Prompts )
   Call NextStep( w, "OFF"     , Prompts )

   GameOver = "True"

End Sub

Sub UpdateSummary( w, Prompts )

   If Len( SummText ) Then

      Print #SummFile, Left$( CustCode & "           ", 11 ) & Left$( ProdCode & "              ", 14 ) & " -- " & SummText

   Else

      If Len( ProdCode ) Then

         Print #SummFile, Left$( CustCode & "           ", 11 ) & Left$( ProdCode & "              ", 14 ) & " -- Successfully ordered (" & OrderNum & "); "

      End If

   End If

   SummText = ""

End Sub

Sub DebugDump( w, Prompts )

   Debug.Print PageLine
   Debug.Print SnapShot
   Debug.Print PageLine
   Debug.Print Action & " ==> " & Reaction & " `" & Prompts( CInt( Reaction ) ) & "`"
   Debug.Print PageLine

   For X = 0 To 99

      If Prompts( X ) = "" Then

         Exit For

      Else

         Debug.Print "Prompts(" & Right$( "  " & Str$( X ), 2 ) & ") = `" & Prompts( X ) & "`"

      End If

   Next X

   Debug.Print PageLine

End Sub
