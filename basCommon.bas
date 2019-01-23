Attribute VB_Name = "basCommon"

Public gfTallyKeyWarningAlreadyShown As Boolean

Public gstrINVVerify1 As String
Public gstrINVVerify2 As String
Public gstrINVVerify3 As String
Public gstrINVVerify4 As String
Public gstrInvVerifyAutoClosed As String

Public gfDebug2013 As Boolean
Public gfLicense As Boolean
Public gintLastUtilityListIndex As Integer
Public gstrStatus As String
Public gPrinterIP As String
Public gFootageTrueMath As Boolean
Public gPrinterPort As Long
Public glngTagAutomatic As Long
Public glngTagmanual As Long
Public gstrAlphaPrefix As String
Public guseHandheldNumbering As Boolean
Public gPRODUCTVERIFICATION As Boolean
Public dblLAYERLENGTH As Double
Public dblshrinkage As Double

Public strScan  As String
Public bScanIt As Boolean
Public glngUserID As Long
Public gstrUserName As String
Public gdblWeekTotal As Double
Public intSystemj As Integer
Public strErrMessage As String
Public gstrUnitDefault As String
Public gfWeekTotal As Boolean
Public gintSMLW As Integer
Public gf7600 As Boolean
Public gf9900 As Boolean
Public gSymbol As Boolean
Public gfHX2 As Boolean
Public gfrmHeight As Long
Public gfrmWidth As Long
Public gDailyDay As Integer
Public gstrIPSM As String
Public gstrModel As String
Public gf99EX As Boolean

Public ginvimportactionflag As Integer
Public gDailyFootage As Double
Public gCullPieces As Double
Public gCullFootage As Double
Public gfBackupNeeded As Boolean
Public gintSMFirst As Integer
Public garyOH() As tOHRecord
Public garyOL() As tOLRecord
'Public garyALL() As tALLRecord delete after 10/1/2009
Public garyCONT() As tCONTRecord
Public garyGrade() As String
Public garyThickness() As String
Public garySpecies() As String
Public garyStatus() As String
Public garyLength() As String
Public gudtLenNew() As tLengthRecord
Public gudtGrade() As tGradeRecord

Public garyOrg() As String
Public garyPA() As tPARecord
Public garyNav_frmGradeLMenu() As String
Public garyUser() As String
Public garyBE() As tBERecord
Public garyLoad() As tLoadRecord
Public garyLoadMem() As tLoadMemRecord
Public garyCTLine() As tCTLineRecord
Public gTallyKey() As String
Public gUnloadCount As Integer
Public garyLR() As tLRRecord
Public garyLRLine() As tLRLineRecord
Public garyMAC() As tMACRecord
Public garyLoc() As tLOCRecord
Public garyProdRun() As tProdRunRecord

Public gudtUtility() As tUtilityRecord

Public gstrServerName As String
Public gStrShipper As String
Public intTimer As Integer
Public gintBackupCountCT As Integer
Public gintBackupCTLine As Integer
Public gintLoadAutoAdvance As Integer
Public glngLoadAutoID As Long


Public Const PROD_STATUS_Green_Inv = 1
Public Const PROD_STATUS_AIR_DRIED_INV2 = 41

Public Const PROD_STATUS_KILN = 2
Public Const PROD_STATUS_TRANSFER_TO_FINAL_GRADING = 3
Public Const PROD_STATUS_KILN_DRIED_INV = 4

Public Const PROD_STATUS_PENDING_SHIPMENT = 5
Public Const PROD_STATUS_SHIPPED = 6
Public Const PROD_STATUS_UNTALLIED = 7
Public Const PROD_STATUS_Green_Inv_SALE = 8
Public Const PROD_STATUS_UNTALLIED_SALE = 9
Public Const PROD_STATUS_MISPLACED_INV = 10
Public Const PROD_STATUS_Breakdown_INV = 12

Public Const PROD_STATUS_GREENLOAD = 30
Public Const PROD_STATUS_KDLOAD = 31

Public Const PROD_STATUS_GREENLR_INV = 32
Public Const PROD_STATUS_KDLR_INV = 33

Public Const PROD_STATUS_GREENLR_CN = 34
Public Const PROD_STATUS_KDLR_CN = 35
Public Const PROD_STATUS_CANTS = 36
Public Const PROD_STATUS_YARDSCAN = 37
Public Const PROD_STATUS_CONSUMED = 40
Public Const PROD_STATUS_INSPECTIONGR = 45
Public Const PROD_STATUS_KILNRESERVED = 11
Public Const PROD_STATUS_STEAMER = 191

Enum TallyKey
    Species
    SpeciesID
    KeyText
    KeyCode
    GradeID
    Grade
    ThicknessID
    thickness
    LoadAID
    GradeHHID
    ThicknessHHID
    Last
End Enum

Enum RecV3MixTags
    MixGroupBundleCount
    MixGroupAID
    MixGroupName
    MixGroupPcsGrade
    MixGroupPcsReject
    MixGroupPcsTotal
    LoadDate
    LoadAID
    BuyerAID
    AllGroups
    Last
End Enum
Public gfDEMO As Boolean
''****************''****************''****************''****************''****************
''****************
''**************** KEYS CHANGED FROM HARD CONSTANTS TO INTEGERS ON 5/18/2018
''**************** Search project for the below sub to find keycodes assigned to variables -- or hit ctrl-End - it's a bottom of this module currently
''**************** Public Sub HHMODEL_KEYCODE_ASSIGNMENT(strModel as String,strKeypad as String)
''****************
''****************''****************''****************''****************''****************

Public keyEnter   As Integer ' Was =13      'ENTER ONLY
Public KeyUp   As Integer ' Was =38        'NOTHING
Public KeyFldExit   As Integer ' Was =149

Public KeyDown   As Integer ' Was =40      'NOTHING
Public keyLeft   As Integer ' Was =37
Public keyRight   As Integer ' Was =39

Public keyAD   As Integer ' Was =65         'HHP A
Public keyKD   As Integer ' Was =74         'HHP J
Public keyCancel   As Integer ' Was =27     'HHP ESC Key
Public KeyHelp   As Integer ' Was =112      'HHP F1
Public KeyHelp_9900   As Integer ' Was =227 'HHP 9900 F1 Key
Public KeyHelp2   As Integer ' Was =45      'HHP INS Key
Public KeyHelp2_99EX As Integer

Public keySave   As Integer ' Was =113      'HHP F2

Public KeyView   As Integer ' Was =114      'HHP F3

Public KeyEdit   As Integer ' Was =115      'HHP F4


Public keyAdd   As Integer ' Was =187    'HHP Blue/SP (+)

Public KeyDelete As Integer    'HHP Blue/Del (-)

Public KeyMax   As Integer ' Was =200   'Place Holder for now
Public KeyMin   As Integer ' Was =200

Public KeyF1   As Integer ' Was =112
Public KeyF2   As Integer ' Was =113
Public KeyF3   As Integer ' Was =114
Public KeyF4   As Integer ' Was =115
Public KeyF5   As Integer ' Was =116
Public KeyF6   As Integer ' Was =96
Public KeyF7   As Integer ' Was =97
Public KeyF8   As Integer ' Was =119
Public KeyF9   As Integer ' Was =120
Public KeyF10   As Integer ' Was =121

Public KeyF1_9900   As Integer ' Was =227
Public KeyF2_9900   As Integer ' Was =228
Public KeyF3_9900   As Integer ' Was =230
Public KeyF4_9900   As Integer ' Was =233
Public KeyF5_9900   As Integer ' Was =234
Public KeyF6_9900   As Integer ' Was =235
Public KeyF7_9900   As Integer ' Was =236
Public KeyF8_9900   As Integer ' Was =237
Public KeyF9_9900   As Integer ' Was =238
Public KeyF10_9900   As Integer ' Was =239

Public KeyScan   As Integer ' Was =42


Public KeySP   As Integer ' Was =32
Public KeyBKSP   As Integer ' Was =8
Public KeyCTRL   As Integer ' Was =17
Public KeySymbolRedDot   As Integer ' Was =126 'This is actual equivalent to the "~"
Public KeyTab   As Integer ' Was =9 'Tab key on the HHP Handhelds
Public KeyComma   As Integer ' Was =188 'This is true on 56 key, next to zero button
Public KeyDecimal   As Integer ' Was =190  'This is true on 56 key, next to zero button
Public KeySemicolon_99EX   As Integer ' Was =186
Public KeyPoundSign_99EX   As Integer ' Was =155
'****************KEYS CHANGED FROM HARD CONSTANTS TO INTEGERS ON 5/18/2018
'*****************KEY CODE VALUE ASSIGNMENT IS NOW IN MODULE - HHMODEL_KEYCODE_ASSIGNMENT(strModel as String,strKeypad as String)

'locations for lumber
Public Const STATUS_AIR_DRIED_INV = 1
Public Const STATUS_KILN = 2
Public Const STATUS_TRANSFER_TO_FINAL_GRADING = 3
Public Const STATUS_KILN_DRIED_INV = 4
Public Const STATUS_PENDING_SHIPMENT = 5
Public Const STATUS_SHIPPED = 6
Public Const STATUS_UNTALLIED = 7
Public Const STATUS_AIR_DRIED_INV_SALE = 8
Public Const STATUS_UNTALLIED_SALE = 9
Public Const STATUS_MISPLACED_INV = 10
Public Const STATUS_GRADEBREAKDOWN = 12

Public fFirstGradeLoad As Boolean
Public fFirstThicknessLoad As Boolean
Public fFirstSpeciesLoad As Boolean
Public fFirstOrgLoad As Boolean
Public fFirstPALoad As Boolean
Public fFirstUserLoad As Boolean
Public fFirstBELoad As Boolean
Public fFirstMacLoad As Boolean
Public fFirstLocLoad As Boolean


Public gfImportCheck As Boolean

'These are for the bundle header "memory"
Public gstrLastRecAID As String
Public gstrLastShiftID As String
Public gstrLastSpecies As String
Public gstrLastBatch As String
Public gstrLastGrade As String
Public gstrLastSurface As String
Public gstrLastLoad As String
Public gstrLastLocation As String
Public gstrLastStatus As String
Public gstrLastThickness As String
Public gstrLastGrader As String
Public gintGradeKeys As Integer
Public gstrLastCustOrgAID As String
Public gstrLastCustOrgID As String
Public gstrLastCustOrg As String
Public gstrLastLoadNumber As String
Public gstrShifID As String
Public gstrHHPSerialNum As String
Public gstrLastLoadBundleID As String
Public gstrLastPDClass As String
Public gstrLastPDClassID As String
Public gstrLastPosition As String
Public gstrLastProdRunAID As String
Public gstrLastLengthUnit As String
Public gstrLastPDW1 As String
Public gstrLastColorAID As String
Public gstrLastProdAID As String
Public gstrLastOrder As String
Public gstrLastMacAID As String
Public gstrLastPDI1 As String
Public gstrLastPDI2 As String
Public gstrLastBundlePrefix As String
Public gstrLastPercent As String
Public gstrLastDimensioned As String
Public gstrLastVendor As String
Public gstrLastCarrier As String
Public gstrLastReceive_FreightType As String
Public gstrLastReceive_FreightAmt As String
Public gstrLastAvgWidth As String
Public gstrLastCalcMethod As String

Public gintGradePercentage As Integer

Public garySecurity() As String
Public gstrPrinterCommMethod As String
Public gintPrinterPort As Long
Public gstrPrinterIP As String

Public gstrPartnerMode As String
Public gstrClient As String

Public gdblShipOveragePercent As Double
Public garyPDB() As String

Public gfSMTallyFormRunning As Boolean

Public Type RecV3Group
    GroupName As String
    GroupThickness As String
    GroupWidth As Double
    GroupGrade As String
    GroupSpecies As String
    GroupPcsOnGrade As Double
    GroupPcsReject As Double
    GroupPcsTotal As Double
        
End Type
Public garyRecV3Group() As RecV3Group

Public Type RecQuality
    RecQualityName As String
    RecQualityValue As String
End Type
Public garyRecQuality() As RecQuality

Public Type Settings
    BundleTodayESTForm As String
    BundleTodayETForm As String
    BundleExportPath As String
    BundleExportFileName As String
    BundleExportType As String
    BundleExportLengthOnly As String
    
    BTComm As Integer
    
    BTGradeDistribute As String
    BTClassReq As String
    BTLockLocAID As String
    BTLockMacAID As String
    BTUseMacAID As String
    BTUseColor As String
    BTUse3Widths As String
    BTUsePercent As String
    BTUsePosition As String
    BTWidthReq As String
    BTStartOnField As String
    BTShiftID As String
    BTStaves As String
    
    
    LRPrefix As String
    LRSEEDValue As Long
    LRSeedValueTag As Long
    ReceiveDetailAutoPrint As String
    ReceiveDetailScanTag As String
    ReceiveDisableAutoSave As String
    ReceiveDefaultLocation As String
    ReceiveX1Label As String
    ReceiveX2Label As String
    ReceiveX3Label As String
    ReceiveQualFields As String
    ReceiveDefaultPO As String
    ReceiveDefaultPWidth As Double
    ReceiveFormVersion As String
    ReceivePrintTicket As String
    ReceiveSkipFields As String
    ReceiveFootageAdjust As String
    
    
    CTPrefix As String
    CTSeedValue As Long
    CTNewLoadNumberAutoonClose As String
    CTDefaultTagsToPrint As String
    CTStartOnAfterLoadMem As String
    CTExportOnSave As String
    CTSoundOnNearLoadTotal As String
    CTSingleLoadOnly As String
    CTTagSuffix As String
    CT_Report_Daily_Version As String
    CT_SMReport_UseRecNumber_As_RecDate As String
    
    GRSeedValueTag As Long
    ADSeedValueTag As Long
    KDSeedVAlueTag As Long
    
    RunIDLookupType As String
    LoginMethod As String
    ReceiveDayLoadOption As String
    
    
    MACAID As String
    
    BlockTallyCustom As String
    BTSerialSettings As String
    BTGPositionDefault As String
    BTDefaultPercent As String
    BTTagFileCount As String
    BTUseWidthForLayerEstimate As String
    
    BTPrintType As String 'Mobile/NetworkFile - either directly to printer via cpcl or writes networkfile with field list
    BTPrintType_FieldList As String 'Comma Separated List of TagNames to replace when printing networkfile instead of tag printer direct example <BundleID>,<THK>,<SPECIES>,etc
    BTTagFilePath As String
    BTTagFileName As String
    
    BundleIDFormat As String
    
    
    CN_LxW_Add10ToLenValue As String
    CN_LxW_MaxAdd10Key As Long
    CN_Override_CommaTrim As String 'disable the comma Key from opening the Grade After Trim / Upgrade sm grade form
    CN_Override_F5Trim As String    'disable the F5 Key from opening the Grade After Trim / Upgrade sm grade form
    CN_Override_F6Receive As String 'Disables the f6 key from going to the receive # box, and enables it as a grade key.
    CN_Override_CtrlRepeat As String 'Disable the Control Key as the repeat key
    CN_LxW_AutoLengthAssign As String ' This Sets the chain tally LXW Export Style to auto select the load based upon a matching length range for the entered board.
    CN_LXW_UseStatusAndRecNum_ForRunNum As String 'When doing a XPLW Chain Tally (Length Width - Export Tally) this combines the Entry in the Loads Receiving# with Status to create the run number, otherwise the run number is just the rec#
    CN_PrintGradeGroupONTag As String 'Print Grade Group instead of individual grades on tag when printing from the chain tally close process
    
    
    CTTagFilePath As String
    CTTagFileName As String
    CN_SpecialGradeTotal_HHAID As String
    CN_SM_ExcludeNetOnSMReport As String
    CN_UseAirDriedStatus_Enabled As String
    CN_UseNewFootageGradeTotals As String
    CN_GradeToPrint As String
    CN_Enable_F_Alt_Keys As String 'Attempt to enable the F1Alt, F2Alt, etc ..settings to work on the Chain Tally SURFACE MEASURE (ONLY AS OF 2/8/17) Entry
    
    CTPrintType As String
    CTPrintType_FieldList As String 'Comma Separated List of TagNames to replace when printing networkfile instead of tag printer direct example <BundleID>,<THK>,<SPECIES>,etc
    
    CompanyName As String
    CompanyAddress1 As String
    CompanyAddress2 As String
    CompanyAddress3 As String 'Added 6/237/17 so most forms/reports won't use this currently
    CompanyPhone As String
    CompanyFax As String
    
    
    BankName As String 'Added 6/237/17 so most forms/reports won't use this currently
    BankAddress1 As String 'Added 6/237/17 so most forms/reports won't use this currently
    BankAddress2 As String 'Added 6/237/17 so most forms/reports won't use this currently
    BankAddress3 As String 'Added 6/237/17 so most forms/reports won't use this currently
    BankPhone As String 'Added 6/237/17 so most forms/reports won't use this currently
    BankID_Numerator As String 'added 8/16 - this numer is a top number over a bottom # like a fraction (Can't do underline font so using 3 lines) 10-2/220 in this case but different by bank
    BankID_Denominator As String 'added 8/16 - this numer is a top number over a bottom # like a fraction (Can't do underline font so using 3 lines) 10-2/220 in this case but different by bank
    BankID_ImagePath As String ' Path for the Fractional ID/Image for the Banks ID# that gets printed by the name/address
    BankID_ImagePixelHeight As String 'Height of Image in Pixels
    BankID_ImagePixelWidth As String 'Width of Image in Pixels
    
    DefaultDimensioned As Integer
    DefaultWidth As String
    DefaultSurface As String
    DefaultThickness As String
    DefaultAvgWidth As String
    DefaultSpecies As String
    
    DebugMode As String
    
    EBSInkLevelIndex As Long
    EBSFontIndex As Long
    EBSDelayCounter_CommPause As Long
    EBS_Comm_LeaveOpen As String
    EBSInkDensityIndex As Long
    EBS_UseSilvaID_GradeGroupOverride_ToPrint As String
    EBS_UseSilvaID_LGradeOverride_ToPrint As String 'Added 12/22/16 by PCC to allow printing the value of the silvatechid field in grades.pdb in place of printing the LGrade (Line Grade) abbreviation as it is in Grades.pdb/tblGrades.HHAID
    EBS_PrinterID_UniqueSendCommand As String
    
    ETKeyF3Location As String
    ETKeyF4Location As String
    ETKeyF5Location As String
    ETKeyF6Location As String
    
    ETKeyF7_Status As String
    ETKeyF8_Status As String
    
    ETLockLocAID As String
    
    ETUsePrefix As String
    ETUseLayers As String
    ETUsePosition As String
    ETUseColor As String
    ETUsePDI1I2 As String
    ETUsePDI2 As String
    
    ETUseKDSeed As String
    
    ETSoftWood As String
    ETKeySound As String
    ETUseRoundMath As String
    ETRunAIDRequired As String
    ETPrintGrossOnTally As String
    ETPrintCount As String
    
    ETLastDate As String
    ET_GradeToPrint As String 'Added to Mimic the CT_GradeToPrint in case ETPrintModule is called in place of ct print module
    ETPrintType As String 'MOBILE or NETWORKFILE - determines if it prints to a tag printer via bluetooth/like zebra p4t in ZPL or writes a network file with a list of comma delimited fields
    ETPrintType_FieldList As String 'Comma Separated List of TagNames to replace when printing networkfile instead of tag printer direct example <BundleID>,<THK>,<SPECIES>,etc
    ETTagFilePath As String
    ETTagFileName As String
    ETTagPrint_BundleTally_NO_CPCL As String
    ETPrintType_TallyFormatVersion As String
    
    
    
    fDebugMode As Boolean
    
    frmLookupMaxWidthValue As String


    KeyF1Alt As Integer
    Keyf2Alt As Integer
    KeyF3Alt As Integer
    KeyF4Alt As Integer
    KeyF5Alt As Integer
    KeyF6Alt As Integer
    
    KeySpaceAsF1 As String
    
    KilnAlphaLimit As String
    KilnNumericLimit As String
    KilnDefaultPosition As String
    KilnDefaultLocation As String
    KilnDefaultStatus As String
    KilnMacPrefix As String
    KilnRunPrefix As String
    KILNFormVersion As String
    KilnPositionAutoIncrement As String
    
    LastRUNID As String
    
    LabelClass As String
    LabelWidth As String
    LabelI1I2 As String
    
    
    LocationValidate As String
    LocationFilterbyUserID As String
      
    MoveStatusDefault As String
    MovePositionFormat As String
    
    
    ORDERAllowOverride As String
    OrderX1Use As String
    OrderX2Use As String
    OrderX3Use As String
    OrderX4Use As String
    OrderX1Caption As String
    OrderX2Caption As String
    OrderX3Caption As String
    OrderX4Caption As String
    OrderMaxMoisture As Double
    OrderMaxLayers As Double
    OrderMaxBundleID As Long
    
    'EBS Printer Setup from Order Allocation Form
    Order_EBSAutoPrint As String
    Order_EBS_UseTemplates As String
    Order_EBS_Setup_1 As String
    Order_EBS_Setup_2 As String
    Order_EBS_Setup_3 As String
    Order_EBS_Setup_4 As String
    
    OrderQualFields As String
    OrderFormVersion As String
    
    Org_IncludeAllVendorsInCarrierList As String 'Includes any TYPE='VENDOR' in the list when hitting F1 when the box is looking for trucker or carrier types (doesn't require orgtypesub=Carrier to show up, shows all vendors on carrier f1 list)
    
    PositionValidate As String
    
    RecPrintPauseCounterValue As String
    
    
    RequiredClass As String
    RequiredSurface As String
    RequiredWidth As String
    RecAvgWidthDeductPercentage As String
    RecExportVersion As String
    RecExportV2PathandName As String
    RecDeleteEnabled As String
    RecTagPrintStyle As String
    
    RecTicketPrint As String
    RecStartOn As String
    RecLoadNumberLocked As String
    RecSwitchViewKey As String
    RecUseLoadNumberAsTagPrefix As String
    RecPrintIndividualTags As String
    
    ReceiveSkipFields_UseToHide As String
    Receive_LayersLabel As String
    Receive_HideLengthRange As String
    
    Reports_PDI1Label  As String
    Reports_Tally_CreateExportFileOnPrint As String
    
    Reports_TallySummaryVersion As String
    RecReport_PrintAsDimensioned As String
    RecReport_TotalX1Numeric As String
    RecReport_TotalX2Numeric As String
    RecPrintCheckEnabled As String
    RecPrintType As String 'Mobile/NetworkFile - either directly to printer via cpcl or writes networkfile with field list
    RecPrintType_FieldList As String 'Comma Separated List of TagNames to replace when printing networkfile instead of tag printer direct example <BundleID>,<THK>,<SPECIES>,etc
    RecTagFilePath As String
    RecTagFileName As String
    
'Receive Form V3 Related Fields And Settings
    ReceiveDefaultIncludeRejects As String 'YES/No Includ Rejects # when Printing Tags 'No=Just show pieces total, not reject count (it's a keep the vendor happy/selling to you reasoning/issue)
    ReceiveDefaultFreightUnit As String 'FLAT/MBF for Identification of how to calculate freight amount
    LRSeedValue_CheckNumber As String 'Next Check Number to Print For Purchasing/Check Printing
    ReceiveV3_Label_BOLID As String 'Allows changing the label on the BOL ID /Ref# Box to use it as something else frmReceiveV3 only (Receive Form Version=V3
    ReceiveV3_Label_PONumber As String 'Allows changing the label on the PO Number Box to use it as something else frmReceiveV3 only (Receive Form Version=V3
    
    ReceiveV3_HideHeaderFields As String  '- Default Value= "Ref#" Used to hide the fields on the V3 form (Header/Load Level fields, not entry
                                        'level/bundle/product level) Must use value / Tag Property Value of the text box minus the txt (txtBOLID for example should be just "BOLID,")
                                        'Separate fields by a comma, even if there is only one field you are hiding
                                        
    ReceiveV3_GotoEditListAfterSave As String    'Causes the system to load the frmList (Add/Edit/Print for Receiving Loads) after clicking the cmdSave or cmdSave_Print button on frmReceiveV3*****
    ReceiveV3_DefaultFreightType As String 'Default value for new recieved loads using v3 of receiving form
    ReceiveLoadReport_FileName As String 'The file name (not path, the files must reside in the IPSM folder on 99ex series handheld or "hard memory" folders on other devices
    ReceiveCheckReport_FileName As String 'The file name (not path, the files must reside in the IPSM folder on 99ex series handheld or "hard memory" folders on other devices
    ReceiveLoadReport_NumberCopies As String 'Sets the default number of copies of the receive load/delivery ticket report to print. Only works in frmReceiveV3 as of 2/13/2017
    ReceiveV3_Auto_NextFieldSelect As String 'Sets the ReceiveV3 to auto move to appropriate entry field when the grade indicates a set thk/width/len dimension or thk/width dimension
        'in some cases will jump to pieces, in some cases will jump to length - for entry.
    ReceiveV3_Price_OnGrade_Max As String ' Sets a max per unit price entry value for On Grade (PCS) pricing during receivev3 load entry
    ReceiveV3_Price_Rejects_Max As String ' Sets a max per unit price entry value for reject pricing during receivev3 load entry
    ReceiveV3_Check_Stub1_TotalRows As String 'Sets the total number of rows when printing checks from receiveV3 - So that it knows when to stop/start the Stub and Start The Check section (check is on middle stub currently 7/31/2017)
    ReceiveV3_Check_Stub1_PrintLimit As String ' Sets the max # of loads that print on the top stub before it screws up the alignment, this is an attempt to autocorret it assuming the difference is less than the # of blank lines being added at the bottom currently.
    
    ReceiveV3_LoadReportVersion As String ''- Changes the load report layout to a completely hardcoded/table/product ID designed layout instead of standard text file layout
    
    Receive_PrintTicket_RequireConfirm As String 'Used for the P4T printer to handle the slow print, for large # of tags on the load, to pause and let the user decide when the printer is done printing bundle tags and ready to print the receive ticket/summary tag
    Receive_ClearX1ForNewBundle As String '- receive v2 - doesn't persist X1 field bundle to bundle, enter each time
    
    SEEDPREFIXTAG As String
    
    Security_Setup_GradeSM_LoadEditLock As String  'Disables the Ability to Edit Open Loads on the Surface Measure Chain Tally Screen (Load Options Button)
    Security_Setup_Thk_Lock As String       'Disables the Ability to Add/Edit Thicknesses on the Handheld
    
    
    ShellCommand1 As String
    ShellCommand1Caption As String
    ShellCommand1KeyCode As String
    ''
    SoundSpeedUseSpecial As String
    SoundSpeed9 As String
    SoundSpeed10 As String
    SoundSpeed40 As String
    SoundSpeed50 As String
    SoundSpeed150 As String
    
    systemM3_FeetToCM_ConversionFactor As String
    systemM3_InchesToCM_ConversionFactor As String
    SystemM3_BFMConversionFactor As String
    
    TopOfFormLock As String
    TopOfFormUse As String
    
    TagPrintDelay_BetweenTags As String
    
    WindowsDeleteShortCuts As String
    
    WIPSeedValue As Long
    
    WirelessType As String
    XferBatch As String
    XferPath As String
    XferAutoTransferTypes As String
    XferLastBackupIndex As String
    
    ZPL_Print_Language As String
    ZPL_Tag_File_Name As String
    InventoryFormVersion As String
    InvFormV2_SkipLoc_SkipEntity As String
    InvFormV2_UseProdQty As String
    
    Receivev3_MixProdTag_FileName As String
    Receive_Check_IncludeFreightType As String
    BTCalcMethod_Options As String ' List of Calculation Methods (Block tally basic only as of 9/2018) that can be selected/rolled through using the F6 key for
        'different calculation methods, tag printing/labeling, etc.  If left blank, includes ONLY BLOCK,PCS
        'Options as of today are BLOCK, LAYERS (estimated footages setup), PCS (same as dimensioned checkbox),Lineal,3Width
    BTCalcMethod_Default As String 'sets the default calculation method that is loaded when opening a new block tally after launching the program
        'If following setting is "YES" then it only happens the first time/defaults it the first time after opening program, if below setting <> YES
        'then it will revert to this every time the form is open/closed
    BTCalcMethod_RememberLast As String ' Yes/No Value that determines if the last calc method is used as the next calc method. If no it reverts back to the ...default setting above each time.
    BTCalcMethod_Layers_Use_MacID As String 'Yes/No Value that determines if the mac id should be used in layer estimated footages/matching
    Report_BundleTally_Version As String 'modified to allow decimal lengths as well as integer lengths 12/12/2019 - VAlid settings are blank or V2 ..v2 has decimal lengths
    
End Type

Public gSettings As Settings



Option Explicit
Public Sub InitializeGlobalVariables()
    Dim strETFilePath_OldSetting As String
    
On Error GoTo ErrorHandler

    Call Utility_Search("Report_BundleTally_Version", "V1", "NAME", True, True)
    gSettings.Report_BundleTally_Version = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveV3_Check_Stub1_PrintLimit", "0", "NAME", True, True)
    gSettings.ReceiveV3_Check_Stub1_PrintLimit = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("BTCalcMethod_Layers_Use_MacID", "NO", "NAME", True, True)
    gSettings.BTCalcMethod_Layers_Use_MacID = tcu(gudtUtility(0).UtilityValueText) ' hot
    
    'This is also in settings below, but not autoadding, so putting up here to to autoadd if not there '
    Call Utility_Search("BTCalcMethod_Options", "Block,Pcs", "NAME", True, True)
    gSettings.BTCalcMethod_Options = tcu(gudtUtility(0).UtilityValueText) ' hot key to switch from onestatus to another (only works on frmbundletally currently) 5/17/2018
    
    Call Utility_Search("BTCalcMethod_Default", "", "NAME", True, True)
    gSettings.BTCalcMethod_Default = tcu(gudtUtility(0).UtilityValueText) ' hot key to switch from onestatus to another (only works on frmbundletally currently) 5/17/2018
    
    Call Utility_Search("BTCalcMethod_RememberLast", "NO", "NAME", True, True)
    gSettings.BTCalcMethod_RememberLast = tcu(gudtUtility(0).UtilityValueText) ' hot key to switch from onestatus to another (only works on frmbundletally currently) 5/17/2018
    
    
    
    Call Utility_Search("ETKeyF7_Status", "", "NAME", True, True)
    gSettings.ETKeyF7_Status = tcu(gudtUtility(0).UtilityValueText) ' hot key to switch from onestatus to another (only works on frmbundletally currently) 5/17/2018
    
    Call Utility_Search("ETKeyF8_Status", "", "NAME", True, True) ' hot key to switch from onestatus to another (only works on frmbundletally currently) 5/17/2018 - enter PSAID of status you want assigned
    gSettings.ETKeyF8_Status = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Receive_Check_IncludeFreightType", "", "NAME", True, True)
    gSettings.Receive_Check_IncludeFreightType = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecPrintPauseCounterValue", "1300", "NAME", True, True)
    gSettings.RecPrintPauseCounterValue = CStr(RR(gudtUtility(0).UtilityValueText))
    
    Call Utility_Search("Receivev3_MixProdTag_FileName", "RecieveTagDetailMultigGradev3_1.LBL", "NAME", True, True)
    gSettings.Receivev3_MixProdTag_FileName = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveV3_LoadReportVersion", "V1", "NAME", True, True)
    gSettings.ReceiveV3_LoadReportVersion = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("InvFormV2_UseProdQty", "YES", "NAME", True, True)
    gSettings.InvFormV2_UseProdQty = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("InvFormV2_SkipLoc_SkipEntity", "YES", "NAME", True, True)
    gSettings.InvFormV2_SkipLoc_SkipEntity = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("InventoryFormVersion", "V1", "NAME", True, True)
    gSettings.InventoryFormVersion = tcu(gudtUtility(0).UtilityValueText)
    
    
    Call Utility_Search("BankID_ImagePixelHeight", "43", "NAME", True, True) '
    gSettings.BankID_ImagePixelHeight = tcu(gudtUtility(0).UtilityValueText)
    If IsNumeric(gSettings.BankID_ImagePixelHeight) = False Then gSettings.BankID_ImagePixelHeight = 43
    
    
    Call Utility_Search("BankID_ImagePixelWidth", "50", "NAME", True, True) '
    gSettings.BankID_ImagePixelWidth = tcu(gudtUtility(0).UtilityValueText)
    If IsNumeric(gSettings.BankID_ImagePixelWidth) = False Then gSettings.BankID_ImagePixelWidth = 50
 
 
    Call Utility_Search("BankID_ImagePath", "\IPSM\BankID.bmp", "NAME", True, True) '
    gSettings.BankID_ImagePath = tcu(gudtUtility(0).UtilityValueText)
    
    

    Call Utility_Search("BankID_Numerator", "", "NAME", True, True) '
    gSettings.BankID_Numerator = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("BankID_Denominator", "", "NAME", True, True) '
    gSettings.BankID_Denominator = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveV3_Check_Stub1_TotalRows", "12", "NAME", True, True) '
    gSettings.ReceiveV3_Check_Stub1_TotalRows = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Receive_ClearX1ForNewBundle", "NO", "NAME", True, True) '
    gSettings.Receive_ClearX1ForNewBundle = tcu(gudtUtility(0).UtilityValueText)
    
    
    Call Utility_Search("Receive_PrintTicket_RequireConfirm", "NO", "NAME", True, True) '
    gSettings.Receive_PrintTicket_RequireConfirm = tcu(gudtUtility(0).UtilityValueText)
    
    
    Call Utility_Search("Security_Setup_GradeSM_LoadEditLock", "NO", "NAME", True, True) '
    gSettings.Security_Setup_GradeSM_LoadEditLock = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Security_Setup_Thk_Lock", "NO", "NAME", True, True) '
    gSettings.Security_Setup_Thk_Lock = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveV3_Price_OnGrade_Max", "99999", "NAME", True, True) 'Added for ReceiveV3-Stella Jones 20170303 by PCC
    gSettings.ReceiveV3_Price_OnGrade_Max = CStr(RR(gudtUtility(0).UtilityValueText))
    If RR(gSettings.ReceiveV3_Price_OnGrade_Max) <= 0 Then gSettings.ReceiveV3_Price_OnGrade_Max = "99999"
    
    Call Utility_Search("ReceiveV3_Price_Rejects_Max", "99999", "NAME", True, True) 'Added for ReceiveV3-Stella Jones 20170303 by PCC
    gSettings.ReceiveV3_Price_Rejects_Max = CStr(RR(gudtUtility(0).UtilityValueText))
    If RR(gSettings.ReceiveV3_Price_Rejects_Max) <= 0 Then gSettings.ReceiveV3_Price_Rejects_Max = "99999"
    
    Call Utility_Search("ReceiveV3_Auto_NextFieldSelect", "YES", "NAME", True, True) 'Added for ReceiveV3-Stella Jones 20170303 by PCC
    gSettings.ReceiveV3_Auto_NextFieldSelect = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveV3_GotoEditListAfterSave", "NO", "NAME", True, True)
    gSettings.ReceiveV3_GotoEditListAfterSave = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveLoadReport_NumberCopies", "1", "NAME", True, True)
    gSettings.ReceiveLoadReport_NumberCopies = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CN_Enable_F_Alt_Keys", "", "NAME", True, True)
    gSettings.CN_Enable_F_Alt_Keys = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveLoadReport_FileName", "", "NAME", True, True)
    gSettings.ReceiveLoadReport_FileName = tcu(gudtUtility(0).UtilityValueText)
    
    
    Call Utility_Search("ReceiveCheckReport_FileName", "", "NAME", True, True)
    gSettings.ReceiveLoadReport_FileName = tcu(gudtUtility(0).UtilityValueText)
    
    '1/17/2017-1/19/2017 - While onsite at SJ for Procurement Project
    Call Utility_Search("ReceiveV3_DefaultFreightType", "FOB-Mill", "NAME", True, True)
    gSettings.ReceiveV3_DefaultFreightType = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Org_IncludeAllVendorsInCarrierList", "YES", "NAME", True, True)
    gSettings.Org_IncludeAllVendorsInCarrierList = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveDefaultIncludeRejects", "NO", "NAME", True, True)
    gSettings.ReceiveDefaultIncludeRejects = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveDefaultFreightUnit", "MBF", "NAME", True, True)
    gSettings.ReceiveDefaultFreightUnit = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("LRSeedValue_CheckNumber", "C1000", "NAME", True, True)
    gSettings.LRSeedValue_CheckNumber = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveV3_Label_BOLID", "Ref #", "NAME", True, True)
    gSettings.ReceiveV3_Label_BOLID = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveDefaultIncludeRejects", "NO", "NAME", True, True)
    gSettings.ReceiveDefaultIncludeRejects = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveV3_HideHeaderFields", "BOLID,PONUMBER,", "NAME", True, True)
    gSettings.ReceiveV3_HideHeaderFields = tcu(gudtUtility(0).UtilityValueText)
    
 
    'End 1/17-1/19 Changes
    
    Call Utility_Search("ETFilePath", "", "NAME", True, False, False)
    
    If SC(gudtUtility(0).UtilityName, "ETFilePath") = True Then
        strETFilePath_OldSetting = gudtUtility(0).UtilityValueText
    End If
    
    'When adding user settings, add a description / usage of the variable in a comment ABOVE the ...Call Utilit..... line as well as your initials/date of change
    Call DeleteUtilityRecordByName("ETFilePath")
    
    Call Utility_Search("EBS_UseSilvaID_LGradeOverride_ToPrint", "NO", "NAME", True, True)
    gSettings.EBS_UseSilvaID_LGradeOverride_ToPrint = tcu(gudtUtility(0).UtilityValueText)
    '**** SEE ABOVE WHEN ADDING NEW ***
    
    Call Utility_Search("ETPrintType", "", "NAME", True, True)
    gSettings.ETPrintType = gudtUtility(0).UtilityValueText
    
    
    
    'End Tally Variable to Add 3 options to what is returned in tag file that includes <BUNDLETALLY> Variable - Added 12/5/2016 - PCC
    Call Utility_Search("ETPrintType_TallyFormatVersion", "V1", "NAME", True, True, False)
    gSettings.ETPrintType_TallyFormatVersion = tcu(gudtUtility(0).UtilityValueText)
    
    'Network File Printing/Variable Field List Option Added Oct 2016-PCC for AHI
    'BlockTally
    
    Call Utility_Search("ETTagPrint_BundleTally_NO_CPCL", "", "NAME", True, True, False)
    gSettings.ETTagPrint_BundleTally_NO_CPCL = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("BTPrintType", "", "NAME", True, True, False)
    gSettings.BTPrintType = tcu(gudtUtility(0).UtilityValueText)
    
    If SC(gSettings.BTPrintType, "MOBILE") = True Then gSettings.BTPrintType = gSettings.ETPrintType
    
    
    Call Utility_Search("BTPrintType_FieldList", "", "NAME", True, True, False)
    gSettings.BTPrintType_FieldList = Trim(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("BTTagFilePath", "", "NAME", True, True, False)
    gSettings.BTTagFilePath = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("BTTagFileName", "", "NAME", True, True, False)
    gSettings.BTTagFileName = tcu(gudtUtility(0).UtilityValueText)
    
    'EndTally
    Call Utility_Search("ETTagFilePath", "", "NAME", True, True, False)
    
    
    gSettings.ETTagFilePath = tcu(gudtUtility(0).UtilityValueText)
    If SC(gSettings.ETTagFilePath, "") = True And SC(strETFilePath_OldSetting, "") = False Then
        gSettings.ETTagFilePath = strETFilePath_OldSetting
    End If
    
    
    Call Utility_Search("ETTagFileName", "", "NAME", True, True, False)
    gSettings.ETTagFileName = tcu(gudtUtility(0).UtilityValueText)
    
    If SC(gSettings.ETTagFileName, "") = True Then 'use the older setting
        Call Utility_Search("TAGFILE-ET", "", "NAME", True, True)
        gSettings.ETTagFileName = gudtUtility(0).UtilityValueText
        Call Utility_Search("ETTagFileName", gSettings.ETTagFileName, "NAME", True, True, True, "UtilityValueText")
        Call Utility_Delete("TagFile-ET", False, "")
        gSettings.ETTagFilePath = gstrPDBPath
        Call Utility_Search("ETTagFilePath", gstrPDBPath, "NAME", True, True, True, "UtilityValueText")
    End If
    
    
    Call Utility_Search("ETPrintType", "MOBILE", "NAME", True, True, False)
    gSettings.ETPrintType = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ETPrintType_FieldList", "", "NAME", True, True, False)
    gSettings.ETPrintType_FieldList = Trim(gudtUtility(0).UtilityValueText)
    
    If SC(gSettings.ETPrintType_FieldList, "MOBILE") = True Then gSettings.ETPrintType_FieldList = ""
        
        
    'ChainTally
    Call Utility_Search("CTPrintType", "MOBILE", "NAME", True, True, False)
    
    gSettings.CTPrintType = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CTPrintType_FieldList", "", "NAME", True, True, False)
    gSettings.CTPrintType_FieldList = Trim(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CTTagFilePath", "", "NAME", True, True, False)
    gSettings.CTTagFilePath = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CTTagFileName", "", "NAME", True, True, False)
    gSettings.CTTagFileName = tcu(gudtUtility(0).UtilityValueText)
    
    'Receiving
    Call Utility_Search("RecPrintType", "MOBILE", "NAME", True, True, False)
    gSettings.RecPrintType = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecPrintType_FieldList", "MOBILE", "NAME", True, True, False)
    gSettings.RecPrintType_FieldList = Trim(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecTagFilePath", "", "NAME", True, True, False)
    gSettings.RecTagFilePath = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecTagFileName", "", "NAME", True, True, False)
    gSettings.RecTagFileName = tcu(gudtUtility(0).UtilityValueText)
    '''End of add for networkfile print version V3 with variable field list by tallytype.
    
    
    Call Utility_Search("ET_GradeToPrint", "GRADE", "NAME", True, True)
    gSettings.ET_GradeToPrint = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CN_GradeToPrint", "GRADE", "NAME", True, True)
    gSettings.CN_GradeToPrint = tcu(gudtUtility(0).UtilityValueText)
    
    'Gets the encoded 8 digits that are sent as part of the communication strings when communicating with the printers so the handheld speaks to a specific printer
    'rather than any that are around ...seems to be required based upon PCC understanding/testing/experience so far.
    gSettings.EBS_PrinterID_UniqueSendCommand = EBSGetPrinterLicense_ReturnUniqueSendCommand()
    
    Call Utility_Search("EBS_UseSilvaID_GradeGroupOverride_ToPrint", "NO", "NAME", True, True)
    gSettings.EBS_UseSilvaID_GradeGroupOverride_ToPrint = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Receive_HideLengthRange", "NO", "NAME", True, True)
    gSettings.Receive_HideLengthRange = tcu(gudtUtility(0).UtilityValueText)
    
    
    Call Utility_Search("Receive_LayersLabel", "Layers", "NAME", True, True)
    gSettings.Receive_LayersLabel = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("DefaultSpecies", "", "NAME", True, True)
    gSettings.DefaultSpecies = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveSkipFields_UseToHide", "NO", "NAME", True, True)
    gSettings.ReceiveSkipFields_UseToHide = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ReceiveSkipFields", "", "NAME", True, True)
    gSettings.ReceiveSkipFields = gudtUtility(0).UtilityValueText
    
    Call Utility_Search("CN_UseNewFootageGradeTotals", "NO", "NAME", True, True, False)
    gSettings.CN_UseNewFootageGradeTotals = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CN_PrintGradeGroupONTag", "NO", "NAME", True, True, False)
    gSettings.CN_PrintGradeGroupONTag = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecPrintCheckEnabled", "NO", "NAME", True, True, False)
    gSettings.RecPrintCheckEnabled = tcu(gudtUtility(0).UtilityValueText)
    
    
    Call Utility_Search("RecReport_TotalX1Numeric", "NO", "NAME", True, True, False)
    gSettings.RecReport_TotalX1Numeric = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecReport_TotalX2Numeric", "NO", "NAME", True, True, False)
    gSettings.RecReport_TotalX2Numeric = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecReport_PrintAsDimensioned", "NO", "NAME", True, True, False)
    gSettings.RecReport_PrintAsDimensioned = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RECEIVEFORMVERSION", "V2", "NAME", True, True, False)
    gSettings.ReceiveFormVersion = tcu(gudtUtility(0).UtilityValueText)
    
        
    Call Utility_Search("Reports_TallySummaryVersion", "V1", "NAME", True, True, False)
    gSettings.Reports_TallySummaryVersion = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CN_UseAirDriedStatus_Enabled", "NO", "NAME", True, True)
    gSettings.CN_UseAirDriedStatus_Enabled = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Reports_Tally_CreateExportFileOnPrint", "NO", "NAME", True, True)
    gSettings.Reports_Tally_CreateExportFileOnPrint = tcu(gudtUtility(0).UtilityValueText)
    
    'For EBSPrinting Defaulting Array Size 4/12/16 - PCC
    ReDim chrArr.ary(100)
    ReDim getArr.ary(100)
    ReDim sendArr.ary(100)
    
    
    Call Utility_Search("EBSInkDensityIndex", "500", "NAME", True, True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
        gSettings.EBSInkDensityIndex = CLng(gudtUtility(0).UtilityValueText)
    Else
        gSettings.EBSInkDensityIndex = 500
    End If
    
    Call Utility_Search("CT_SMReport_UseRecNumber_As_RecDate", "NO", "NAME", True, True)
    gSettings.CT_SMReport_UseRecNumber_As_RecDate = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("KEYF5ALT", "-1", "NAME", True, True)
    If SC(gudtUtility(0).UtilityName, "KEYF5ALT") = True Then
        If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
            gSettings.KeyF5Alt = CInt(AV(gudtUtility(0).UtilityValueText))
        Else
            gSettings.KeyF5Alt = -1
        End If
    End If
    
    Call Utility_Search("KEYF6ALT", "-1", "NAME", True, True)
    If SC(gudtUtility(0).UtilityName, "KEYF6ALT") = True Then
        If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
            gSettings.KeyF6Alt = CInt(AV(gudtUtility(0).UtilityValueText))
        Else
            gSettings.KeyF6Alt = -1
        End If
    End If
    
            
    Call Utility_Search("Reports_PDI1Label", "I1", "NAME", True, True)
    gSettings.Reports_PDI1Label = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CT_Report_Daily_Version", "V1", "NAME", True, True)
    gSettings.CT_Report_Daily_Version = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("EBSDelayCounter_CommPause", "NO", "NAME", True, True)
    gSettings.EBS_Comm_LeaveOpen = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("EBSDelayCounter_CommPause", "900", "NAME", True, True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then gudtUtility(0).UtilityValueText = "0"
    gSettings.EBSDelayCounter_CommPause = CLng(gudtUtility(0).UtilityValueText)
    If gSettings.EBSDelayCounter_CommPause < 500 Then gSettings.EBSDelayCounter_CommPause = 900
    
    Call Utility_Search("DebugMode", "", "NAME", True, True)
    gSettings.DebugMode = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("EBSInkLevelIndex", "0", "NAME", True, True, False)
    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
        gSettings.EBSInkLevelIndex = CLng(gudtUtility(0).UtilityValueText)
    Else
        gSettings.EBSInkLevelIndex = 1
    End If
    
    Call Utility_Search("EBSFontIndex", "0", "NAME", True, True, False)
    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
        gSettings.EBSFontIndex = CLng(gudtUtility(0).UtilityValueText)
    Else
        gSettings.EBSFontIndex = 0
    End If
    
    Call Utility_Search("Order_EBS_UseTemplates", "NO", "NAME", True, True, False)
    gSettings.Order_EBS_UseTemplates = tcu(gudtUtility(0).UtilityValueText)
    
    
    'EBS Template Printing Options from Order Allocation
    Call Utility_Search("Order_EBS_Setup_1", "<BUNDLEID>" & "~*~" & "<ORDERID>" & "~*~" & "<SPECIES>  <THK>  <GRADE>  <STATUS>" & "~*~" & "<ORDERID>", "NAME", True, True, False)
    gSettings.Order_EBS_Setup_1 = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Order_EBS_Setup_2", "", "NAME", True, True, False)
    gSettings.Order_EBS_Setup_2 = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Order_EBS_Setup_3", "", "NAME", True, True, False)
    gSettings.Order_EBS_Setup_3 = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Order_EBS_Setup_4", "", "NAME", True, True, False)
    gSettings.Order_EBS_Setup_4 = tcu(gudtUtility(0).UtilityValueText)
    
    
    
    'End 4/12/16 PCC
    Call Utility_Search("OrderFormVersion", "V2", "NAME", True, True, False)
    gSettings.OrderFormVersion = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("Order_EBSAutoPrint", "NO", "NAME", True, True, False)
    gSettings.Order_EBSAutoPrint = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CN_LXW_UseStatusAndRecNum_ForRunNum", "YES", "NAME", True, True, False)
    gSettings.CN_LXW_UseStatusAndRecNum_ForRunNum = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CN_SM_ExcludeNetOnSMReport", "NO", "NAME", True, True, False)
    gSettings.CN_SM_ExcludeNetOnSMReport = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("CN_SpecialGradeTotal_HHAID", "", "NAME", True, True, False)
    gSettings.CN_SpecialGradeTotal_HHAID = tc(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ZPL_Print_Language", "CPCL", "NAME", True, True, False)
    gSettings.ZPL_Print_Language = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ZPL_Tag_File_Name", "ChainTag.lbl", "NAME", True, True, False)
    gSettings.ZPL_Tag_File_Name = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("BTComm", "4", "NAME", True, True, False)
    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
        gSettings.BTComm = CInt(gudtUtility(0).UtilityValueText)
    Else
        gSettings.BTComm = 4
    End If
    
    
    Call Utility_Search("ZPL_Print_Language", "CPCL", "NAME", True, True, False)
    gSettings.ZPL_Print_Language = tcu(gudtUtility(0).UtilityValueText)
    
    
    'Special Sound on Keypress options
    Call Utility_Search("BTSerialSettings", "57600,N,8,1", "NAME", True, True, False)
    gSettings.BTSerialSettings = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("TagPrintDelay_BetweenTags", "500", "NAME", True, True)
    gSettings.TagPrintDelay_BetweenTags = tcu(gudtUtility(0).UtilityValueText)
    
    If IsNumeric(gSettings.TagPrintDelay_BetweenTags) = True Then
        If CLng(gSettings.TagPrintDelay_BetweenTags) > 0 Then
            'do nothing
        Else
            gSettings.TagPrintDelay_BetweenTags = "500"
        End If
    Else
        gSettings.TagPrintDelay_BetweenTags = "500"
    End If
    
    Call Utility_Search("BTUseWidthForLayerEstimate", "NO", "NAME", True, True)
    gSettings.BTUseWidthForLayerEstimate = tcu(gudtUtility(0).UtilityValueText)
    
    
    Call Utility_Search("SoundSpeedUseSpecial", "NO", "NAME", True, True)
    gSettings.SoundSpeedUseSpecial = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("SoundSpeed9", "0", "NAME", True, True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then gudtUtility(0).UtilityValueText = "0"
    gSettings.SoundSpeed9 = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("SoundSpeed10", "0", "NAME", True, True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then gudtUtility(0).UtilityValueText = "0"
    gSettings.SoundSpeed10 = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("SoundSpeed40", "0", "NAME", True, True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then gudtUtility(0).UtilityValueText = "0"
    gSettings.SoundSpeed40 = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("SoundSpeed50", "0", "NAME", True, True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then gudtUtility(0).UtilityValueText = "0"
    gSettings.SoundSpeed50 = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("SoundSpeed150", "0", "NAME", True, True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then gudtUtility(0).UtilityValueText = "0"
    gSettings.SoundSpeed150 = tcu(gudtUtility(0).UtilityValueText)
    '''End of special sound settings
    
    '''***********Conversion Factors for BTMetric
    Call Utility_Search("systemM3_FeetToCM_ConversionFactor", "30.48", "NAME", True, True)
    gSettings.systemM3_FeetToCM_ConversionFactor = tcu(gudtUtility(0).UtilityValueText)
    
    If IsNumeric(gSettings.systemM3_FeetToCM_ConversionFactor) = False Then gSettings.systemM3_FeetToCM_ConversionFactor = "0"
    If CDbl(gSettings.systemM3_FeetToCM_ConversionFactor) = 0 Then
        gSettings.systemM3_FeetToCM_ConversionFactor = "30.48"
    End If
    
    Call Utility_Search("systemM3_InchesToCM_ConversionFactor", "2.54", "NAME", True, True)
    gSettings.systemM3_InchesToCM_ConversionFactor = tcu(gudtUtility(0).UtilityValueText)
    If IsNumeric(gSettings.systemM3_InchesToCM_ConversionFactor) = False Then gSettings.systemM3_InchesToCM_ConversionFactor = "0"
    If CDbl(gSettings.systemM3_InchesToCM_ConversionFactor) = 0 Then
        gSettings.systemM3_InchesToCM_ConversionFactor = "2.54"
    End If
    
    Call Utility_Search("SystemM3_BFMConversionFactor", "0.0283168465921", "NAME", True, True)
    gSettings.SystemM3_BFMConversionFactor = tcu(gudtUtility(0).UtilityValueText)
    If IsNumeric(gSettings.SystemM3_BFMConversionFactor) = False Then gSettings.SystemM3_BFMConversionFactor = "0"
    If CDbl(gSettings.SystemM3_BFMConversionFactor) = 0 Then
        gSettings.SystemM3_BFMConversionFactor = "0.0283168465921"
    End If
    '''***********END OF Conversion Factors for BTMetric
    
    Call Utility_Search("CN_LxW_AutoLengthAssign", "NO", "NAME", True, True)
    gSettings.CN_LxW_AutoLengthAssign = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("BTTagFileCount", "2", "NAME", True, True)
    gSettings.BTTagFileCount = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("ETPrintCount", "1", "NAME", True, True)
    gSettings.ETPrintCount = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecPrintIndividualTags", "NO", "NAME", True, True)
    gSettings.RecPrintIndividualTags = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecUseLoadNumberAsTagPrefix", "NO", "NAME", True, True, False)
    gSettings.RecUseLoadNumberAsTagPrefix = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecLoadNumberLocked", "YES", "NAME", True, True)
    gSettings.RecLoadNumberLocked = tcu(gudtUtility(0).UtilityValueText)
    
    Call Utility_Search("RecSwitchViewKey", "0", "NAME", True, True)
    gSettings.RecSwitchViewKey = tcu(gudtUtility(0).UtilityValueText)
    If IsNumeric(gSettings.RecSwitchViewKey) = False Then
        gSettings.RecSwitchViewKey = "0"
    End If

    Exit Sub
ErrorHandler:
    MsgBox "Error in InitializeGlobalVariables: " & Err.Number & "-" & Err.Description
    Exit Sub
End Sub

Public Sub InitializeGlobalVariablesV2()
    ReDim gudtUtility(0)
    Dim lngUtilCount As Long
    Dim I As Integer
    
On Error GoTo ErrorHandler
    OpenUtilityDatabase
    If dbUtility = 0 Then
        MsgBox "Could not open utility database, contact support!"
        End
    End If
    
    lngUtilCount = PDBNumRecords(dbUtility)
    ReDim gudtUtility(lngUtilCount - 1)
    
    PDBBulkRead dbUtility, lngUtilCount, VarPtr(gudtUtility(I))
    
    For I = 0 To UBound(gudtUtility)
        If SC(gudtUtility(I).UtilityName, "RecPrintPauseCounterValue") = True Then
            gSettings.RecPrintPauseCounterValue = tcu(gudtUtility(I).UtilityValueText)
            If IsNumeric(gSettings.RecPrintPauseCounterValue) = True Then
                If CLng(gSettings.RecPrintPauseCounterValue) >= 1 Then
                    'just continue
                Else
                    gSettings.RecPrintPauseCounterValue = "10"
                End If
            Else
                gSettings.RecPrintPauseCounterValue = "10"
            End If
        End If
        
        
        If SC(gudtUtility(I).UtilityName, "ETPrintGrossOnTally") = True Then
            If SC(gudtUtility(I).UtilityValueText, "YES") = True Then
                gSettings.ETPrintGrossOnTally = "1"
            Else
                gSettings.ETPrintGrossOnTally = "0"
            End If
        End If
        
        If SC(gudtUtility(I).UtilityName, "RecPrintIndividualTags") = True Then
            gSettings.RecPrintIndividualTags = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "RecUseLoadNumberAsTagPrefix") = True Then
            gSettings.RecUseLoadNumberAsTagPrefix = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "RecSwitchViewKey") = True Then
            gSettings.RecSwitchViewKey = tcu(gudtUtility(I).UtilityValueText)
            If IsNumeric(gSettings.RecSwitchViewKey) = False Then gSettings.RecSwitchViewKey = "0"
        End If
        
        
        If SC(gudtUtility(I).UtilityName, "RecLoadNumberLocked") = True Then
            gSettings.RecLoadNumberLocked = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "XferLastBackupIndex") = True Then
            gSettings.XferLastBackupIndex = tcu(gudtUtility(I).UtilityValueText)
            If IsNumeric(gSettings.XferLastBackupIndex) = False Then gSettings.XferLastBackupIndex = "1"
        End If
        
        If SC(gudtUtility(I).UtilityName, "BTGPositionDefault") = True Then
            gSettings.BTGPositionDefault = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "RecStartOn") = True Then
            gSettings.RecStartOn = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "RecTicketPrint") = True Then
            gSettings.RecTicketPrint = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "ETRunAIDRequired") = True Then
            gSettings.ETRunAIDRequired = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "RecTagPrintStyle") = True Then
            gSettings.RecTagPrintStyle = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "XferAutoTransferTypes") = True Then
            gSettings.XferAutoTransferTypes = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "WirelessType") = True Then
            gSettings.WirelessType = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        
        If SC(gudtUtility(I).UtilityName, "BTStaves") = True Then
            gSettings.BTStaves = tcu(gudtUtility(I).UtilityValueText)
        End If
    
    
        If SC(gudtUtility(I).UtilityName, "CTTagSuffix") = True Then
            gSettings.CTTagSuffix = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "CTPrintType") = True Then
            gSettings.CTPrintType = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "CTNetworkTagFilePath") = True Then
            If SC(gSettings.CTTagFilePath, "") = True Then
                gSettings.CTTagFilePath = tcu(gudtUtility(0).UtilityValueText)
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "CTNetworkTagFileName") = True Then
            If SC(gSettings.CTTagFileName, "") = True Then
                gSettings.CTTagFileName = tcu(gudtUtility(0).UtilityValueText)
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "RecDeleteEnabled") = True Then
            gSettings.RecDeleteEnabled = tcu(gudtUtility(I).UtilityValueText)
        End If



        If SC(gudtUtility(0).UtilityName, "BTSerialSettings") = True Then
        
            gSettings.BTSerialSettings = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        
        If SC(gudtUtility(I).UtilityName, "CTSingleLoadOnly") = True Then
            gSettings.CTSingleLoadOnly = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "CTSoundOnNearLoadTotal") = True Then
            gSettings.CTSoundOnNearLoadTotal = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "CTExportOnSave") = True Then
        
            gSettings.CTExportOnSave = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "CTStartOnAfterLoadMem") = True Then
            gSettings.CTStartOnAfterLoadMem = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "CTDefaultTagsToPrint") = True Then
            gSettings.CTDefaultTagsToPrint = Trim(UCase(gudtUtility(I).UtilityValueText))
        End If
        
        If IsNumeric(gSettings.CTDefaultTagsToPrint) = False Then gSettings.CTDefaultTagsToPrint = "0"
        

        If SC(gudtUtility(I).UtilityName, "CTNewLoadNumberAutoonClose") = True Then
            gSettings.CTNewLoadNumberAutoonClose = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETSoftWood") = True Then
            gSettings.ETSoftWood = Trim(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BundleExportLengthOnly") = True Then
            gSettings.BundleExportLengthOnly = Trim(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BundleExportPath") = True Then
            gSettings.BundleExportPath = Trim(gudtUtility(I).UtilityValueText)
            If Right(Trim(gSettings.BundleExportPath), 1) <> "\" Then gSettings.BundleExportPath = Trim(gSettings.BundleExportPath) & "\"
        End If

        If SC(gudtUtility(I).UtilityName, "BundleExportFileName") = True Then
            gSettings.BundleExportFileName = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BundleExportType") = True Then
            gSettings.BundleExportType = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUseRoundMath") = True Then
            gSettings.ETUseRoundMath = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RecExportV2PathandName") = True Then
            gSettings.RecExportV2PathandName = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "RECExportVersion") = True Then
            gSettings.RecExportVersion = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "CN_Override_CtrlRepeat") = True Then
            gSettings.CN_Override_CtrlRepeat = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "CN_Override_CommaTrim") = True Then
            gSettings.CN_Override_CommaTrim = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "CN_Override_F5Trim") = True Then
            gSettings.CN_Override_F5Trim = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "CN_Override_F6Receive") = True Then
            gSettings.CN_Override_F6Receive = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "CN_LxW_MaxAdd10Key") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.CN_LxW_MaxAdd10Key = CLng(gudtUtility(I).UtilityValueText)
            Else
                gSettings.CN_LxW_MaxAdd10Key = -1
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "CN_LxW_Add10ToLenValue") = True Then
            gSettings.CN_LxW_Add10ToLenValue = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "LOCATIONFILTERBYUSERID") = True Then
            gSettings.LocationFilterbyUserID = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "LOCATIONFILTERBYUSERID") = True Then
            gSettings.LocationFilterbyUserID = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "LOCATIONFILTERBYUSERID") = True Then
            gSettings.LocationFilterbyUserID = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BTSHIFTID") = True Then
            gSettings.BTShiftID = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BTSTARTONFIELD") = True Then
            gSettings.BTStartOnField = tcu(gudtUtility(I).UtilityValueText)

        End If

        If SC(gudtUtility(I).UtilityName, "DEFAULTAVGWIDTH") = True Then
            gSettings.DefaultAvgWidth = tcu(gudtUtility(I).UtilityValueText)

            If IsNumeric(gSettings.DefaultAvgWidth) = False Then
                gSettings.DefaultAvgWidth = "38"
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "BTGRADEDISTRIBUTE") = True Then
            gSettings.BTGradeDistribute = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "WindowsDeleteShortCuts") = True Then
            gSettings.WindowsDeleteShortCuts = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "KEYSPACEASF1") = True Then
            gSettings.KeySpaceAsF1 = tcu(gudtUtility(I).UtilityValueText)

            If tcu(gSettings.KeySpaceAsF1) = "" Then gSettings.KeySpaceAsF1 = "YES"
        End If

        If SC(gudtUtility(I).UtilityName, "ETKEYSOUND") = True Then
            gSettings.ETKeySound = tcu(gudtUtility(I).UtilityValueText)
            If gSettings.ETKeySound = "" Then gSettings.ETKeySound = "YES"
        End If
        If SC(gudtUtility(I).UtilityName, "SHELLCOMMAND1") = True Then
            gSettings.ShellCommand1 = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "SHELLCOMMAND1CAPTION") = True Then
            gSettings.ShellCommand1Caption = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, UCase("ShellCommand1KeyCode")) = True Then
            gSettings.ShellCommand1KeyCode = tcu(gudtUtility(I).UtilityValueText)

            If IsNumeric(gSettings.ShellCommand1KeyCode) = False Then
                gSettings.ShellCommand1KeyCode = "0"
            Else
                gSettings.ShellCommand1KeyCode = tcu(gudtUtility(I).UtilityValueText)
            End If
        End If
        If SC(gudtUtility(I).UtilityName, UCase("RecAvgWidthDeductPercentage")) = True Then
            gSettings.RecAvgWidthDeductPercentage = tcu(gudtUtility(I).UtilityValueText)

            If IsNumeric(gSettings.RecAvgWidthDeductPercentage) = False Then
                gSettings.RecAvgWidthDeductPercentage = "0"
            Else
                gSettings.RecAvgWidthDeductPercentage = tcu(gudtUtility(I).UtilityValueText)
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "LOCATIONVALIDATE") = True Then
            gSettings.LocationValidate = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "MOVEPOSITIONFORMAT") = True Then
            gSettings.MovePositionFormat = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEFOOTAGEADJUST") = True Then
            gSettings.ReceiveFootageAdjust = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVESKIPFIELDS") = True Then
            gSettings.ReceiveSkipFields = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "KILNRUNPREFIX") = True Then
            gSettings.KilnRunPrefix = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEDEFAULTPO") = True Then
            gSettings.ReceiveDefaultPO = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEDEFAULTPWIDTH") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.ReceiveDefaultPWidth = CDbl(gudtUtility(I).UtilityValueText)
            Else
                gSettings.ReceiveDefaultPWidth = 39
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "MOVESTATUSDEFAULT") = True Then
            gSettings.MoveStatusDefault = tcu(gudtUtility(I).UtilityValueText)

        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEDEFAULTLOCATION") = True Then
            gSettings.ReceiveDefaultLocation = tcu(gudtUtility(I).UtilityValueText)
        End If


        If SC(gudtUtility(I).UtilityName, "RECEIVEX1LABEL") = True Then
            gSettings.ReceiveX1Label = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEX2LABEL") = True Then
            gSettings.ReceiveX2Label = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEX3LABEL") = True Then
            gSettings.ReceiveX3Label = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEQUALFIELDS") = True Then
            gSettings.ReceiveQualFields = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERMAXBUNDLEID") = True Then

            If IsNumeric(gudtUtility(I).UtilityValueText) = False Then
                gSettings.OrderMaxBundleID = 10
            Else
                gSettings.OrderMaxBundleID = CDbl(gudtUtility(I).UtilityValueText)
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "ORDERMAXLAYERS") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = False Then
                gSettings.OrderMaxLayers = 99
            Else
                gSettings.OrderMaxLayers = CDbl(gudtUtility(I).UtilityValueText)
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERMAXMOISTURE") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = False Then
                gSettings.OrderMaxMoisture = 99
            Else
                gSettings.OrderMaxMoisture = CDbl(gudtUtility(I).UtilityValueText)
            End If

        End If

        If SC(gudtUtility(I).UtilityName, "COMPANYNAME") = True Then
            gSettings.CompanyName = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "COMPANYADDRESS1") = True Then
            gSettings.CompanyAddress1 = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "COMPANYADDRESS2") = True Then
            gSettings.CompanyAddress2 = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "COMPANYADDRESS3") = True Then
            gSettings.CompanyAddress3 = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "COMPANYPHONE") = True Then
            gSettings.CompanyPhone = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "COMPANYFAX") = True Then
            gSettings.CompanyFax = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        'Start Bank settings added 6/227/17 by PCC
        If SC(gudtUtility(I).UtilityName, "BankName") = True Then
            gSettings.BankName = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BankAddress1") = True Then
            gSettings.BankAddress1 = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BankAddress2") = True Then
            gSettings.BankAddress2 = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BankAddress3") = True Then
            gSettings.BankAddress3 = tcu(gudtUtility(I).UtilityValueText)
        End If
        
        If SC(gudtUtility(I).UtilityName, "BankPhone") = True Then
            gSettings.BankPhone = tcu(gudtUtility(I).UtilityValueText)
        End If
        'End Bank settings added 6/227/17 by PCC
        
        If SC(gudtUtility(I).UtilityName, "KILNMACPREFIX") = True Then
            gSettings.KilnMacPrefix = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUSEKDSEED") = True Then
            gSettings.ETUseKDSeed = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERX1USE") = True Then
            gSettings.OrderX1Use = tcu(gudtUtility(I).UtilityValueText)
            gSettings.OrderX1Caption = tcu(gudtUtility(I).UtilityValue2Text)
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERX2USE") = True Then
            gSettings.OrderX2Use = tcu(gudtUtility(I).UtilityValueText)
            gSettings.OrderX2Caption = tcu(gudtUtility(I).UtilityValue2Text)
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERX3USE") = True Then
            gSettings.OrderX3Use = tcu(gudtUtility(I).UtilityValueText)
            gSettings.OrderX3Caption = tcu(gudtUtility(I).UtilityValue2Text)
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERX4USE") = True Then
            gSettings.OrderX4Use = tcu(gudtUtility(I).UtilityValueText)
            gSettings.OrderX4Caption = tcu(gudtUtility(I).UtilityValue2Text)
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERQUALFIELDS") = True Then
            gSettings.OrderQualFields = tcu(gudtUtility(I).UtilityValueText)
        End If

        If SC(gudtUtility(I).UtilityName, "DEBUGMODE") = True Then

            If SC(gudtUtility(I).UtilityValueText, "YES") Then
                gSettings.fDebugMode = True
            Else
                gSettings.fDebugMode = False
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "TOPOFFORMUSE") = True Then
            gSettings.TopOfFormUse = tcu(gudtUtility(I).UtilityValueText)
            If SC(gSettings.TopOfFormUse, "YES") = True Then
                gSettings.TopOfFormUse = "1"
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEDISABLEAUTOSAVE") = True Then
            gSettings.ReceiveDisableAutoSave = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "BUNDLEIDFORMAT") = True Then
            gSettings.BundleIDFormat = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "POSITIONVALIDATE") = True Then
            gSettings.PositionValidate = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "KILNDEFAULTLOCATION") = True Then
            gSettings.KilnDefaultLocation = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "KILNDEFAULTSTATUS") = True Then
            gSettings.KilnDefaultStatus = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "KILNDEFAULTPOSITION") = True Then
            gSettings.KilnDefaultPosition = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "KILNALPHALIMIT") = True Then
            gSettings.KilnAlphaLimit = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "KILNNUMERICLIMIT") = True Then
            gSettings.KilnNumericLimit = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUSEPREFIX") = True Then
            gSettings.ETUsePrefix = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUSEPOSITION") = True Then
            gSettings.ETUsePosition = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUSELAYERS") = True Then
            gSettings.ETUseLayers = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUSECOLOR") = True Then
            gSettings.ETUseColor = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUSEPDI1I2") = True Then
            gSettings.ETUsePDI1I2 = tcu(gudtUtility(I).UtilityValueText)
        End If
        If SC(gudtUtility(I).UtilityName, "ETUSEPDI2") = True Then
            gSettings.ETUsePDI2 = tcu(gudtUtility(I).UtilityValueText)
        End If


        If SC(gudtUtility(I).UtilityName, "BACKUPCOUNT-CT") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gintBackupCountCT = CLng(gudtUtility(I).UtilityValueText)
            Else
                gintBackupCountCT = 0
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "KEYF1ALT") = True Then
            gSettings.KeyF1Alt = CInt(AV(gudtUtility(I).UtilityValueText))
        End If
        If SC(gudtUtility(I).UtilityName, "KEYF2ALT") = True Then
            gSettings.Keyf2Alt = CInt(AV(gudtUtility(I).UtilityValueText))
        End If
        If SC(gudtUtility(I).UtilityName, "KEYF3ALT") = True Then
            gSettings.KeyF3Alt = CInt(AV(gudtUtility(I).UtilityValueText))
        End If
        
        If SC(gudtUtility(I).UtilityName, "KEYF4ALT") = True Then
            gSettings.KeyF4Alt = CInt(AV(gudtUtility(I).UtilityValueText))
        End If
        If SC(gudtUtility(I).UtilityName, "TOPOFFORMLOCK") = True Then
            gSettings.TopOfFormLock = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "KILNPOSITIONAUTOINCREMENT") = True Then
            gSettings.KilnPositionAutoIncrement = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "REQUIREDCLASS") = True Then
            gSettings.RequiredClass = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "REQUIREDWIDTH") = True Then
            gSettings.RequiredWidth = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "REQUIREDSURFACE") = True Then
            gSettings.RequiredSurface = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "XFERBATCH") = True Then
            gSettings.XferBatch = gudtUtility(I).UtilityValueText

            If IsNumeric(gSettings.XferBatch) = False Then
                gSettings.XferBatch = CStr(1000)
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "XFERPATH") = True Then
            gSettings.XferPath = gudtUtility(I).UtilityValueText
        End If

        If SC(gudtUtility(I).UtilityName, "BTDEFAULTPERCENT") = True Then
   '''MsgBox "Delete Disabled DEFAULT PERCENT"
        
        'Call Utility_Delete("DEFAULTPERCENT", True, "DBUpdate20130711_DefPercentChange")

            gSettings.BTDefaultPercent = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTUSEPOSITION") = True Then
            gSettings.BTUsePosition = gudtUtility(I).UtilityValueText
        End If

        If SC(gudtUtility(I).UtilityName, "BTUSEPERCENT") = True Then
            gSettings.BTUsePercent = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTUSE3WIDTHS") = True Then
            gSettings.BTUse3Widths = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTCLASSREQ") = True Then
            gSettings.BTClassReq = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTWIDTHREQ") = True Then
            gSettings.BTWidthReq = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTUSECOLOR") = True Then
            gSettings.BTUseColor = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTUSEMACAID") = True Then
            gSettings.BTUseMacAID = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BUNDLETODAYETFORM") = True Then
            gSettings.BundleTodayETForm = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "ORDERALLOWOVERRIDE") = True Then
            gSettings.ORDERAllowOverride = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "KILNFORMVERSION") = True Then
            gSettings.KILNFormVersion = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTLOCKMACAID") = True Then
            gSettings.BTLockMacAID = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "ETLOCKLOCAID") = True Then
            gSettings.ETLockLocAID = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BTLOCKLOCAID") = True Then
            gSettings.BTLockLocAID = gudtUtility(I).UtilityValueText
        End If

        If SC(gudtUtility(I).UtilityName, "WIPSEEDVALUE") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = False Then
                gSettings.WIPSeedValue = 1
            Else
                gSettings.WIPSeedValue = CLng(gudtUtility(I).UtilityValueText)
            End If
        End If


        If SC(gudtUtility(I).UtilityName, "DEFAULTSURFACE") = True Then
            gstrLastPDW1 = gSettings.DefaultSurface
            gSettings.DefaultSurface = gudtUtility(I).UtilityValueText
            gstrLastPDW1 = gSettings.DefaultSurface
        End If
        If SC(gudtUtility(I).UtilityName, "DEFAULTWIDTH") = True Then
            gSettings.DefaultWidth = gudtUtility(I).UtilityValueText
            gstrLastPDW1 = gSettings.DefaultWidth
        End If

        If SC(gudtUtility(I).UtilityName, "FRMLOOKUPMAXWIDTHVALUE") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = False Then
                gSettings.frmLookupMaxWidthValue = 1000
            Else
                gSettings.frmLookupMaxWidthValue = CStr(CLng(gudtUtility(I).UtilityValueText))
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "LABELCLASS") = True Then
            gSettings.LabelClass = gudtUtility(I).UtilityValueText

            If gSettings.LabelClass = "" Then gSettings.LabelClass = "Class"
        End If
        If SC(gudtUtility(I).UtilityName, "LABELWIDTH") = True Then
            gSettings.LabelWidth = gudtUtility(I).UtilityValueText
            If gSettings.LabelWidth = "" Then gSettings.LabelWidth = "Width"
        End If
        If SC(gudtUtility(I).UtilityName, "LABELI1I2") = True Then
            gSettings.LabelI1I2 = gudtUtility(I).UtilityValueText
            If gSettings.LabelI1I2 = "" Then gSettings.LabelI1I2 = "I1/I2"
        End If



        If SC(gudtUtility(I).UtilityName, "RUNIDLOOKUPTYPE") = True Then
            gSettings.RunIDLookupType = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "SEEDPREFIXCT") = True Then
            gSettings.CTPrefix = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "SEEDVALUECT") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.CTSeedValue = CLng(gudtUtility(I).UtilityValueText)
            Else
                gSettings.CTSeedValue = 1
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "SEEDPREFIXLR") = True Then
            gSettings.LRPrefix = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "SEEDVALUELR") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.LRSEEDValue = CLng(gudtUtility(I).UtilityValueText)
            Else
                gSettings.LRSEEDValue = 1000
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "SEEDVALUETAG-LR") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.LRSeedValueTag = CLng(gudtUtility(I).UtilityValueText)
            Else
                gSettings.LRSeedValueTag = 250000
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "SEEDPREFIXTAG") = True Then
            gSettings.SEEDPREFIXTAG = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "SEEDVALUETAG-K") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.KDSeedVAlueTag = CLng(gudtUtility(I).UtilityValueText)
            Else
                gSettings.KDSeedVAlueTag = 800000
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "SEEDVALUETAG-G") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.GRSeedValueTag = CLng(gudtUtility(I).UtilityValueText)
            Else
                gSettings.GRSeedValueTag = 400000
            End If

            If gSettings.GRSeedValueTag < 100 Then gSettings.GRSeedValueTag = 101
        End If

        If SC(gudtUtility(I).UtilityName, "SEEDVALUETAG-K") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gSettings.ADSeedValueTag = CLng(gudtUtility(I).UtilityValueText)
            Else
                gSettings.ADSeedValueTag = 600000
            End If
        End If

        If SC(gudtUtility(I).UtilityName, "BTCOMMRECEIVE") = True Then
            glngBTCommReceive = gudtUtility(I).UtilityValueLong
        End If
        
        If SC(gudtUtility(I).UtilityName, "RECEIVEESTIMATEMETHOD") = True Then
            gstrReceiveEstimateMethod = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "SHIPOVERAGEPERCENT") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gdblShipOveragePercent = CDbl(gudtUtility(I).UtilityValueText)
            Else
                gdblShipOveragePercent = 0
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "ENDTALLYMAXWIDTH") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gintEndTallyMaxWidth = CInt(gudtUtility(I).UtilityValueText)
            Else
                gintEndTallyMaxWidth = 29
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "ENDTALLYTWOKEYMAX") = True Then
            If IsNumeric(gudtUtility(I).UtilityValueText) = True Then
                gintEndTallyTwoKeyMax = CInt(gudtUtility(I).UtilityValueText)
            Else
                gintEndTallyTwoKeyMax = 2
            End If
        End If
        If SC(gudtUtility(I).UtilityName, "ETKEYF3LOCATION") = True Then
            gSettings.ETKeyF3Location = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "ETKEYF4LOCATION") = True Then
            gSettings.ETKeyF4Location = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "ETKEYF5LOCATION") = True Then
            gSettings.ETKeyF5Location = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "ETKEYF6LOCATION") = True Then
            gSettings.ETKeyF6Location = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEDETAILAUTOPRINT") = True Then
            gSettings.ReceiveDetailAutoPrint = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "LOGINMETHOD") = True Then
            gSettings.LoginMethod = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEDAYLOADOPTION") = True Then
            gSettings.ReceiveDayLoadOption = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "RECEIVEDETAILSCANTAG") = True Then
            gSettings.ReceiveDetailScanTag = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "DEFAULTTHICKNESS") = True Then
            gSettings.DefaultThickness = gudtUtility(I).UtilityValueText
            gstrLastThickness = gSettings.DefaultThickness
        End If
        If SC(gudtUtility(I).UtilityName, "MACAID") = True Then
            gSettings.MACAID = gudtUtility(I).UtilityValueText
        End If
        If SC(gudtUtility(I).UtilityName, "BLOCKTALLYCUSTOM") = True Then
            gSettings.BlockTallyCustom = Trim(UCase(gudtUtility(I).UtilityValueText))

        End If

        If SC(gudtUtility(I).UtilityName, "LASTRUNID") = True Then
            gSettings.LastRUNID = Trim(UCase(gudtUtility(I).UtilityValueText))
            gstrLastProdRunAID = gSettings.LastRUNID

        End If
        If SC(gudtUtility(I).UtilityName, "DEFAULTDIMENSIONED") = True Then

            If IsNumeric(Trim(UCase(gudtUtility(I).UtilityValueText))) = True Then
                gSettings.DefaultDimensioned = CInt(Trim(UCase(gudtUtility(I).UtilityValueText)))
            Else
                gSettings.DefaultDimensioned = 0
            End If

        End If




        If SC(gudtUtility(I).UtilityName, "BUNDLETODAYESTFORM") = True Then
            gSettings.BundleTodayESTForm = Trim(UCase(gudtUtility(I).UtilityValueText))

        End If
    Next
    
    
    If gSettings.KeyF1Alt = 0 Then gSettings.KeyF1Alt = -1
    If gSettings.Keyf2Alt = 0 Then gSettings.Keyf2Alt = -1
    If gSettings.KeyF3Alt = 0 Then gSettings.KeyF3Alt = -1
    If gSettings.KeyF4Alt = 0 Then gSettings.KeyF4Alt = -1
    
    'Clean UP the ones we renamed/replaced 11/11/16
    Call Utility_Delete("CTNetworkTagFileName", False, "")
    Call Utility_Delete("CTNetworkTagFilePath", False, "")
    Call Utility_Search("CTTagFileName", gSettings.CTTagFileName, "NAME", True, True, True, "UtilityValueText")
    Call Utility_Search("CTTagFilePath", gSettings.CTTagFilePath, "NAME", True, True, True, "UtilityValueText")

    
    CloseUtilityDatabase
    dbUtility = 0
    ReDim gudtUtility(I)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during setting initialization, Contact Technical Support at 888.520.1951 (" & Err.Number & "-" & Err.Description & ")"
    Exit Sub
    
End Sub
Public Sub InitializeGlobalVariablesCustom()
    ReDim gudtUtility(0)
    Dim lngUtilCount As Long
    Dim I As Integer
    
On Error GoTo ErrorHandler
    ReDim garyRecV3Group(2)
    
    I = 0
    garyRecV3Group(I).GroupName = "7"""
    garyRecV3Group(I).GroupThickness = "7"
    garyRecV3Group(I).GroupGrade = "XT"
    garyRecV3Group(I).GroupSpecies = "ANY"
    garyRecV3Group(I).GroupPcsReject = 0
    garyRecV3Group(I).GroupPcsOnGrade = 0
    garyRecV3Group(I).GroupPcsTotal = 0
    
    I = I + 1
    garyRecV3Group(I).GroupName = "6"""
    garyRecV3Group(I).GroupThickness = "6"
    garyRecV3Group(I).GroupGrade = "XT"
    garyRecV3Group(I).GroupSpecies = "ANY"
    garyRecV3Group(I).GroupPcsReject = 0
    garyRecV3Group(I).GroupPcsOnGrade = 0
    garyRecV3Group(I).GroupPcsTotal = 0
    
    I = I + 1
    garyRecV3Group(I).GroupName = "ST"""
    garyRecV3Group(I).GroupThickness = "ANY"
    garyRecV3Group(I).GroupGrade = "ST"
    garyRecV3Group(I).GroupSpecies = "ANY"
    garyRecV3Group(I).GroupPcsReject = 0
    garyRecV3Group(I).GroupPcsOnGrade = 0
    garyRecV3Group(I).GroupPcsTotal = 0
    

    Exit Sub
    
ErrorHandler:
    MsgBox "Error during InitializeGlobalVariablesCustom , Contact Technical Support at 888.520.1951 (" & Err.Number & "-" & Err.Description & ")"
    Exit Sub
    
End Sub
'''*** OLD VERSION
'''Public Sub InitializeGlobalVariables()
Public Function Receive_GetRecQualityArray(udtLR As tLRRecord, strType As String) As Boolean
    Dim I As Integer
    Dim strSplit As String
    Dim strSplitVal As String
    
On Error GoTo ErrorHandler
    ReDim garyRecQuality(0)
    
    ReDim pstrSplit.ary(0)
    
    If SC(strType, "RECEIVESCAN") = True Then
        strSplit = gSettings.ReceiveQualFields
        strSplitVal = udtLR.PDClass
    ElseIf SC(strType, "RECEIVEV3") = True Then
        strSplit = gSettings.ReceiveQualFields
        strSplitVal = udtLR.TagID
    End If
    
    If InStr(strSplit, ",") > 0 Then
        Call AppForge_Split(strSplit, pstrSplit, ",")
    Else
        ReDim pstrSplit.ary(0)
        pstrSplit.ary(0) = tcu(strSplit)
    End If
    
    ReDim garyRecQuality(UBound(pstrSplit.ary))
    
    
    For I = 0 To UBound(pstrSplit.ary)
        garyRecQuality(I).RecQualityName = pstrSplit.ary(I)
    Next
    
    'Now get the actual values entered for this load, and split and load to the data array for printing/showing/etc
    ReDim pstrSplit.ary(0)
    
    If InStr(strSplitVal, ",") > 0 Then
        Call AppForge_Split(strSplitVal, pstrSplit, ",")
    Else
        pstrSplit.ary(0) = tcu(strSplitVal)
    End If
    
    For I = 0 To UBound(pstrSplit.ary)
        If I <= UBound(garyRecQuality) Then 'just make sure not goofed up and more data than possible defects (if defect list changed while a saved load had used previous list this could happen)
            garyRecQuality(I).RecQualityValue = pstrSplit.ary(I)
        Else
            'do nothing, no place in array to put value...there is a setup/save issue or something
        End If
    Next
    
    Receive_GetRecQualityArray = True
    
    Exit Function
ErrorHandler:
    Receive_GetRecQualityArray = False
    MsgBox "Error In Receive_GetRecQualityArray " & Err.Number & "-" & Err.Description

End Function
Public Function ReceiveV3_GetRecGroupTotals(lngLRID As Long, strLRAID As String) As Boolean

    Dim strWidthAID As String, strGradeAID As String
    Dim dblPcsOnGrade As Double, dblPcsReject As Double
    Dim I As Long
    Dim J As Long
    Dim fMatchFound As Boolean
    
    For I = 0 To UBound(garyRecV3Group)
        garyRecV3Group(I).GroupPcsOnGrade = 0
        garyRecV3Group(I).GroupPcsReject = 0
        garyRecV3Group(I).GroupPcsTotal = 0
    Next
    
On Error GoTo ErrorHandler
    For I = 0 To UBound(garyLRLine)
        If SC(garyLRLine(I).LRLineXString1, "REC-PROD") = True Then
            fMatchFound = False
            For J = 0 To UBound(garyRecV3Group)
                Debug.Print garyRecV3Group(J).GroupGrade & " - LRLine=  " & garyLRLine(I).Grade
                
                If SC(garyRecV3Group(J).GroupGrade, garyLRLine(I).Grade) = True Or SC(garyRecV3Group(J).GroupGrade, Left(garyLRLine(I).Grade, _
                        Len(garyRecV3Group(J).GroupGrade))) = True Or SC(garyRecV3Group(J).GroupGrade, "ANY") = True Then
                    If SC(garyRecV3Group(J).GroupSpecies, garyLRLine(I).Species) = True Or SC(garyRecV3Group(J).GroupSpecies, "ANY") = True Then
                        If SC(garyRecV3Group(J).GroupThickness, garyLRLine(I).thickness) = True Or SC(garyRecV3Group(J).GroupThickness, "ANY") = True Then
                            'It's a match to the group, update the totals
                            If SC(garyLRLine(I).Grade, "XT-IG") = True Or SC(garyLRLine(I).Grade, "XT-CU") = True Then
                                '''No Change Here garyRecV3Group(J).GroupPcsOnGrade = garyRecV3Group(J).GroupPcsOnGrade + garyLRLine(I).LRLineXLong1
                                garyRecV3Group(J).GroupPcsReject = garyRecV3Group(J).GroupPcsReject + garyLRLine(I).LRLineXLong1 'Add IG Counts to the reject Counts
                                garyRecV3Group(J).GroupPcsTotal = garyRecV3Group(J).GroupPcsTotal + garyLRLine(I).LRLineXLong1 'Add IG Counts Total Counts ..shouldn't be any rejects counts on an ig they are rejects
                            Else
                                garyRecV3Group(J).GroupPcsOnGrade = garyRecV3Group(J).GroupPcsOnGrade + garyLRLine(I).LRLineXLong1
                                garyRecV3Group(J).GroupPcsReject = garyRecV3Group(J).GroupPcsReject + garyLRLine(I).LRLineXLong2
                                garyRecV3Group(J).GroupPcsTotal = garyRecV3Group(J).GroupPcsTotal + garyLRLine(I).LRLineXLong1 + garyLRLine(I).LRLineXLong2
                            End If
                            
                            fMatchFound = True
                            Exit For
                        End If
                    End If
                End If
            Next
            
            If fMatchFound = False Then
                MsgBox "Product Group for " & garyLRLine(I).Species & " " & garyLRLine(I).Grade & " " & garyLRLine(I).thickness & "X" & garyLRLine(I).AvgWidthBoard & "X" & garyLRLine(I).AvgLenBoard
            End If
        End If
    Next
        
    'Grade Summary Groups Completed for Load# " & strLRAID
    
    ReceiveV3_GetRecGroupTotals = True
    Exit Function
ErrorHandler:
    ReceiveV3_GetRecGroupTotals = False
    MsgBox "Error in ReceiveV3_GetRecGroupTotals " & Err.Number & "-" & Err.Description
    Exit Function
    
End Function
Private Sub LoadQualRows(grdData As AFGrid, strType As String, strQualityValues As String)
    Dim I As Integer
    Dim strSplit As String

On Error GoTo ErrorHandler

    grdData.Rows = 0
    grdData.Clear
    
    ReDim pstrSplit.ary(0)
    
    If strType = "ORDER" Then
        strSplit = gSettings.OrderQualFields
    ElseIf SC(strType, "RECEIVESCAN") = True Or SC(strType, "RECEIVEV3") = True Then
        strSplit = gSettings.ReceiveQualFields
    End If
    
    If InStr(strSplit, ",") > 0 Then
        Call AppForge_Split(strSplit, pstrSplit, ",")
    Else
        ReDim pstrSplit.ary(0)
        pstrSplit.ary(0) = tcu(strSplit)
    End If
    
    For I = 0 To UBound(pstrSplit.ary)
        grdData.AddItem pstrSplit.ary(I) & vbTab & ""
    Next
    
    ReDim pstrSplit.ary(0)
    strSplit = strQualityValues
    If InStr(strQualityValues, ",") > 0 Then
        Call AppForge_Split(strSplit, pstrSplit, ",")
    Else
        pstrSplit.ary(0) = tcu(strSplit)
    End If
    
    For I = 0 To UBound(pstrSplit.ary)
        grdData.TextMatrix(I, 1) = pstrSplit.ary(I)
    Next
    
    Exit Sub
ErrorHandler:
    MsgBox "Error In LoadQualRows_Type=" & strType & " Error # " & Err.Number & "-" & Err.Description
End Sub

'''On Error GoTo ErrorHandler
    
    
    
    
'    Call Utility_Search("CTTagSuffix", "", "NAME", True, True)
'    gSettings.CTTagSuffix = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTPrintType", "NETWORKFILE", "NAME", True, True)
'    gSettings.CTPrintType = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTTagFilePath", "\\PCName\Temp\", "NAME", True, True)
'    gSettings.CTTagFilePath = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTTagFileName", "TAGBUNDLE.tag", "NAME", True, True)
'    gSettings.CTTagFileName = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RecDeleteEnabled", "NO", "NAME", True, True)
'    gSettings.RecDeleteEnabled = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("BTSerialSettings", "", "NAME", True, True)
'    gSettings.BTSerialSettings = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTSingleLoadOnly", "NO", "NAME", True, True)
'    gSettings.CTSingleLoadOnly = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTSoundOnNearLoadTotal", "NO", "NAME", True, True)
'    gSettings.CTSoundOnNearLoadTotal = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTExportOnSave", "NO", "NAME", True, True)
'    gSettings.CTExportOnSave = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTStartOnAfterLoadMem", "LoadID", "NAME", True, True)
'    gSettings.CTStartOnAfterLoadMem = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CTDefaultTagsToPrint", "0", "NAME", True, True)
'    gSettings.CTDefaultTagsToPrint = Trim(UCase(gudtUtility(0).UtilityValueText))
'
'    If IsNumeric(gSettings.CTDefaultTagsToPrint) = False Then
'        gSettings.CTDefaultTagsToPrint = "0"
'    End If
'
'    Call Utility_Search("CTNewLoadNumberAutoonClose", "NO", "NAME", True, True)
'    gSettings.CTNewLoadNumberAutoonClose = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETSoftWood", "NO", "NAME", True, True)
'    gSettings.ETSoftWood = Trim(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("BundleExportLengthOnly", "NO", "NAME", True, True)
'    gSettings.BundleExportLengthOnly = Trim(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("BundleExportPath", "\", "NAME", True, True)
'    gSettings.BundleExportPath = Trim(gudtUtility(0).UtilityValueText)
'    If Right(Trim(gSettings.BundleExportPath), 1) <> "\" Then gSettings.BundleExportPath = Trim(gSettings.BundleExportPath) & "\"
'
'    Call Utility_Search("BundleExportFileName", "HWT_PRD_<BUNDLEID>_<DATE>.txt", "NAME", True, True)
'    gSettings.BundleExportFileName = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("BundleExportType", "REPORT", "NAME", True, True)
'    gSettings.BundleExportType = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUseRoundMath", "", "NAME", True, True)
'    gSettings.ETUseRoundMath = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RecExportV2PathandName", "\ReceiveExportV2.txt", "NAME", True, True)
'    gSettings.RecExportV2PathandName = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("RECExportVersion", "", "NAME", True, True)
'    gSettings.RecExportVersion = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CN_Override_CtrlRepeat", "NO", "NAME", True, True)
'    gSettings.CN_Override_CtrlRepeat = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CN_Override_CommaTrim", "NO", "NAME", True, True)
'    gSettings.CN_Override_CommaTrim = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CN_Override_F5Trim", "NO", "NAME", True, True)
'    gSettings.CN_Override_F5Trim = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("CN_Override_F6Receive", "NO", "NAME", True, True)
'    gSettings.CN_Override_F6Receive = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("CN_LxW_MaxAdd10Key", "-1", "NAME", True, True)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.CN_LxW_MaxAdd10Key = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.CN_LxW_MaxAdd10Key = -1
'    End If
'
'
'    Call Utility_Search("CN_LxW_Add10ToLenValue", "YES", "NAME", True, True)
'    gSettings.CN_LxW_Add10ToLenValue = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("LOCATIONFILTERBYUSERID", "NO", "NAME", True, True)
'    gSettings.LocationFilterbyUserID = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("LOCATIONFILTERBYUSERID", "NO", "NAME", True, True)
'    gSettings.LocationFilterbyUserID = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("LOCATIONFILTERBYUSERID", "NO", "NAME", True, True)
'    gSettings.LocationFilterbyUserID = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("BTSHIFTID", "", "NAME", True, True, False, "VALUE")
'    gSettings.BTShiftID = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("BTSTARTONFIELD", "THICKNESS", "NAME", True, True, False, "VALUE")
'    gSettings.BTStartOnField = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("DEFAULTAVGWIDTH", "38", "NAME", True, True, False, "VALUE")
'    gSettings.DefaultAvgWidth = tcu(gudtUtility(0).UtilityValueText)
'
'    If IsNumeric(gSettings.DefaultAvgWidth) = False Then
'        gSettings.DefaultAvgWidth = "38"
'    End If
'
'
'    Call Utility_Search("BTGRADEDISTRIBUTE", "NO", "NAME", True, True, False, "VALUE")
'    gSettings.BTGradeDistribute = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("WindowsDeleteShortCuts", "NO", "NAME", True, True, False, "VALUE")
'    gSettings.WindowsDeleteShortCuts = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KEYSPACEASF1", "", "NAME", True, True, False, "VALUE")
'    gSettings.KeySpaceAsF1 = tcu(gudtUtility(0).UtilityValueText)
'
'    If tcu(gSettings.KeySpaceAsF1) = "" Then gSettings.KeySpaceAsF1 = "YES"
'
'
'    Call Utility_Search("ETKEYSOUND", "", "NAME", True, True, False, "VALUE")
'    gSettings.ETKeySound = tcu(gudtUtility(0).UtilityValueText)
'    If gSettings.ETKeySound = "" Then gSettings.ETKeySound = "YES"
'
'    Call Utility_Search("SHELLCOMMAND1", "", "NAME", True, True, False, "VALUE")
'    gSettings.ShellCommand1 = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("SHELLCOMMAND1CAPTION", "", "NAME", True, True, False, "VALUE")
'    gSettings.ShellCommand1Caption = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search(UCase("ShellCommand1KeyCode"), "", "NAME", True, True, False, "VALUE")
'    gSettings.ShellCommand1KeyCode = tcu(gudtUtility(0).UtilityValueText)
'
'    If IsNumeric(gSettings.ShellCommand1KeyCode) = False Then
'        gSettings.ShellCommand1KeyCode = "0"
'    Else
'        gSettings.ShellCommand1KeyCode = tcu(gudtUtility(0).UtilityValueText)
'    End If
'
'    Call Utility_Search(UCase("RecAvgWidthDeductPercentage"), "", "NAME", True, True, False, "VALUE")
'    gSettings.RecAvgWidthDeductPercentage = tcu(gudtUtility(0).UtilityValueText)
'
'    If IsNumeric(gSettings.RecAvgWidthDeductPercentage) = False Then
'        gSettings.RecAvgWidthDeductPercentage = "0"
'    Else
'        gSettings.RecAvgWidthDeductPercentage = tcu(gudtUtility(0).UtilityValueText)
'    End If
'
'    Call Utility_Search("LOCATIONVALIDATE", "", "NAME", True, True, False, "VALUE")
'    gSettings.LocationValidate = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("MOVEPOSITIONFORMAT", "", "NAME", True, True, False, "VALUE")
'    gSettings.MovePositionFormat = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEFOOTAGEADJUST", "NO", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveFootageAdjust = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVESKIPFIELDS", "", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveSkipFields = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KILNRUNPREFIX", "", "NAME", True, True, False, "VALUE")
'    gSettings.KilnRunPrefix = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEDEFAULTPO", "VERBAL", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveDefaultPO = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEDEFAULTPWIDTH", "39", "NAME", True, True, False, "VALUE")
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.ReceiveDefaultPWidth = CDbl(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.ReceiveDefaultPWidth = 39
'    End If
'
'
'    Call Utility_Search("MOVESTATUSDEFAULT", "", "NAME", True, True, False, "VALUE")
'    gSettings.MoveStatusDefault = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("RECEIVEDEFAULTLOCATION", "", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveDefaultLocation = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEFORMVERSION", "V1", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveFormVersion = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("RECEIVEX1LABEL", "", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveX1Label = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEX2LABEL", "", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveX2Label = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEX3LABEL", "Class", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveX3Label = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEQUALFIELDS", "", "NAME", True, True, False, "VALUE")
'    gSettings.ReceiveQualFields = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ORDERMAXBUNDLEID", "10", "NAME", True, True, False, "VALUE")
'
'    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then
'        gSettings.OrderMaxBundleID = 10
'    Else
'        gSettings.OrderMaxBundleID = CDbl(gudtUtility(0).UtilityValueText)
'    End If
'
'
'    Call Utility_Search("ORDERMAXLAYERS", "99", "NAME", True, True, False, "VALUE")
'    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then
'        gSettings.OrderMaxLayers = 99
'    Else
'        gSettings.OrderMaxLayers = CDbl(gudtUtility(0).UtilityValueText)
'    End If
'
'    Call Utility_Search("ORDERMAXMOISTURE", "99", "NAME", True, True, False, "VALUE")
'    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then
'        gSettings.OrderMaxMoisture = 99
'    Else
'        gSettings.OrderMaxMoisture = CDbl(gudtUtility(0).UtilityValueText)
'    End If
'
'
'
'    Call Utility_Search("COMPANYNAME", "", "NAME", True, True, False, "VALUE")
'    gSettings.CompanyName = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("COMPANYADDRESS1", "", "NAME", True, True, False, "VALUE")
'    gSettings.CompanyAddress1 = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("COMPANYADDRESS2", "", "NAME", True, True, False, "VALUE")
'    gSettings.CompanyAddress2 = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("COMPANYPHONE", "", "NAME", True, True, False, "VALUE")
'    gSettings.CompanyPhone = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("COMPANYFAX", "", "NAME", True, True, False, "VALUE")
'    gSettings.CompanyFax = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KILNMACPREFIX", "", "NAME", True, True, False, "VALUE")
'    gSettings.KilnMacPrefix = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUSEKDSEED", "YES", "NAME", True, True, False, "VALUE")
'    gSettings.ETUseKDSeed = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ORDERX1USE", "", "NAME", True, True, False)
'    gSettings.OrderX1Use = tcu(gudtUtility(0).UtilityValueText)
'    gSettings.OrderX1Caption = tcu(gudtUtility(0).UtilityValue2Text)
'
'    Call Utility_Search("ORDERX2USE", "", "NAME", True, True, False)
'    gSettings.OrderX2Use = tcu(gudtUtility(0).UtilityValueText)
'    gSettings.OrderX2Caption = tcu(gudtUtility(0).UtilityValue2Text)
'
'    Call Utility_Search("ORDERX3USE", "", "NAME", True, True, False)
'    gSettings.OrderX3Use = tcu(gudtUtility(0).UtilityValueText)
'    gSettings.OrderX3Caption = tcu(gudtUtility(0).UtilityValue2Text)
'
'    Call Utility_Search("ORDERX4USE", "", "NAME", True, True, False)
'    gSettings.OrderX4Use = tcu(gudtUtility(0).UtilityValueText)
'    gSettings.OrderX4Caption = tcu(gudtUtility(0).UtilityValue2Text)
'
'    Call Utility_Search("ORDERQUALFIELDS", "Inspector,Packaging,Moisture,Finish,Banding,End Paint,Stencil,Tallies,Thickness,Order Spec,# Bundles,Cust,End Checks,Assemble,Release,RLDate,Sales,SLDate,Arriv,In,Out,Load", "NAME", True, True, False)
'    gSettings.OrderQualFields = tcu(gudtUtility(0).UtilityValueText)
'
'
'    Call Utility_Search("DEBUGMODE", "", "NAME", True, True, False)
'
'    If SC(gudtUtility(0).UtilityValueText, "YES") Then
'        gSettings.fDebugMode = True
'    Else
'        gSettings.fDebugMode = False
'    End If
'
'
'    Call Utility_Search("TOPOFFORMUSE", "", "NAME", True, True, False)
'    gSettings.TopOfFormUse = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("RECEIVEDISABLEAUTOSAVE", "", "NAME", True, True, False)
'    gSettings.ReceiveDisableAutoSave = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("BUNDLEIDFORMAT", "", "NAME", True, True, False)
'    gSettings.BundleIDFormat = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("POSITIONVALIDATE", "NO", "NAME", True, True, False)
'    gSettings.PositionValidate = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KILNDEFAULTLOCATION", "", "NAME", True, True, False)
'    gSettings.KilnDefaultLocation = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KILNDEFAULTSTATUS", "", "NAME", True, True, False)
'    gSettings.KilnDefaultStatus = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KILNDEFAULTPOSITION", "", "NAME", True, True, False)
'    gSettings.KilnDefaultPosition = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KILNALPHALIMIT", "Z", "NAME", True, True, False)
'    gSettings.KilnAlphaLimit = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("KILNNUMERICLIMIT", "8", "NAME", True, True, False)
'    gSettings.KilnNumericLimit = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUSEPREFIX", "YES", "NAME", True, True, False)
'    gSettings.ETUsePrefix = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUSEPOSITION", "YES", "NAME", True, True, False)
'    gSettings.ETUsePosition = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUSELAYERS", "YES", "NAME", True, True, False)
'    gSettings.ETUseLayers = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUSECOLOR", "YES", "NAME", True, True, False)
'    gSettings.ETUseColor = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUSEPDI1I2", "YES", "NAME", True, True, False)
'    gSettings.ETUsePDI1I2 = tcu(gudtUtility(0).UtilityValueText)
'
'    Call Utility_Search("ETUSEPDI2", "YES", "NAME", True, True, False)
'    gSettings.ETUsePDI2 = tcu(gudtUtility(0).UtilityValueText)
'
'    '
'
'
'    Call Utility_Search("BACKUPCOUNT-CT", CStr(gintBackupCountCT), "NAME", True, True, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gintBackupCountCT = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gintBackupCountCT = 0
'    End If
'
'    Call Utility_Search("KEYF1ALT", "196", "NAME", True, True, False)
'    gSettings.KeyF1Alt = CInt(AV(gudtUtility(0).UtilityValueText))
'
'    Call Utility_Search("KEYF2ALT", "197", "NAME", True, True, False)
'    gSettings.Keyf2Alt = CInt(AV(gudtUtility(0).UtilityValueText))
'
'    Call Utility_Search("KEYF3ALT", "198", "NAME", True, True, False)
'    gSettings.KeyF3Alt = CInt(AV(gudtUtility(0).UtilityValueText))
'
'    Call Utility_Search("KEYF4ALT", "199", "NAME", True, True, False)
'    gSettings.KeyF4Alt = CInt(AV(gudtUtility(0).UtilityValueText))
'
'    Call Utility_Search("TOPOFFORMLOCK", "", "NAME", True, True, False)
'    gSettings.TopOfFormLock = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("KILNPOSITIONAUTOINCREMENT", "", "NAME", True, True, False)
'    gSettings.KilnPositionAutoIncrement = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("REQUIREDCLASS", "", "NAME", True, True, False)
'    gSettings.RequiredClass = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("REQUIREDWIDTH", "", "NAME", True, True, False)
'    gSettings.RequiredWidth = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("REQUIREDSURFACE", "", "NAME", True, True, False)
'    gSettings.RequiredSurface = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("XFERBATCH", "", "NAME", True, True, False)
'    gSettings.XferBatch = gudtUtility(0).UtilityValueText
'
'    If IsNumeric(gSettings.XferBatch) = False Then
'        gSettings.XferBatch = CStr(1000)
'    End If
'
'    Call Utility_Search("XFERPATH", "", "NAME", True, True, False)
'    gSettings.XferPath = gudtUtility(0).UtilityValueText
'
'    Call Utility_Delete("DEFAULTPERCENT", True, "DBUpdate20130711_DefPercentChange")
'
'    Call Utility_Search("BTDEFAULTPERCENT", "", "NAME", True, True, False)
'    gSettings.BTDefaultPercent = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTUSEPOSITION", "NO", "NAME", True, True, False)
'    gSettings.BTUsePosition = gudtUtility(0).UtilityValueText
'
'
'    Call Utility_Search("BTUSEPERCENT", "", "NAME", True, True, False)
'    gSettings.BTUsePercent = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTUSE3WIDTHS", "", "NAME", True, True, False)
'    gSettings.BTUse3Widths = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTCLASSREQ", "", "NAME", True, True, False)
'    gSettings.BTClassReq = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTWIDTHREQ", "", "NAME", True, True, False)
'    gSettings.BTWidthReq = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTUSECOLOR", "", "NAME", True, True, False)
'    gSettings.BTUseColor = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTUSEMACAID", "", "NAME", True, True, False)
'    gSettings.BTUseMacAID = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BUNDLETODAYETFORM", "", "NAME", True, True, False)
'    gSettings.BundleTodayETForm = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("ORDERALLOWOVERRIDE", "", "NAME", True, True, False)
'    gSettings.ORDERAllowOverride = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("KILNFORMVERSION", "", "NAME", True, True, False)
'    gSettings.KILNFormVersion = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTLOCKMACAID", "", "NAME", True, True, False)
'    gSettings.BTLockMacAID = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("ETLOCKLOCAID", "", "NAME", True, True, False)
'    gSettings.ETLockLocAID = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BTLOCKLOCAID", "", "NAME", True, True, False)
'    gSettings.BTLockLocAID = gudtUtility(0).UtilityValueText
'
'
'    Call Utility_Search("WIPSEEDVALUE", "", "NAME", True, True, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then
'        gSettings.WIPSeedValue = 1
'    Else
'        gSettings.WIPSeedValue = CLng(gudtUtility(0).UtilityValueText)
'    End If
'
'    gstrLastPDW1 = gSettings.DefaultSurface
'
'    Call Utility_Search("DEFAULTSURFACE", "", "NAME", True, True, False)
'    gSettings.DefaultSurface = gudtUtility(0).UtilityValueText
'    gstrLastPDW1 = gSettings.DefaultSurface
'
'    Call Utility_Search("DEFAULTWIDTH", "", "NAME", True, True, False)
'    gSettings.DefaultWidth = gudtUtility(0).UtilityValueText
'    gstrLastPDW1 = gSettings.DefaultWidth
'
'
'    Call Utility_Search("FRMLOOKUPMAXWIDTHVALUE", "", "NAME", True, True, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then
'        gSettings.frmLookupMaxWidthValue = 1000
'    Else
'        gSettings.frmLookupMaxWidthValue = CStr(CLng(gudtUtility(0).UtilityValueText))
'    End If
'
'
'    Call Utility_Search("LABELCLASS", "", "NAME", True, True, False)
'    gSettings.LabelClass = gudtUtility(0).UtilityValueText
'    If gSettings.LabelClass = "" Then gSettings.LabelClass = "Class"
'
'    Call Utility_Search("LABELWIDTH", "", "NAME", True, True, False)
'    gSettings.LabelWidth = gudtUtility(0).UtilityValueText
'    If gSettings.LabelWidth = "" Then gSettings.LabelWidth = "Width"
'
'    Call Utility_Search("LABELI1I2", "", "NAME", True, True, False)
'    gSettings.LabelI1I2 = gudtUtility(0).UtilityValueText
'    If gSettings.LabelI1I2 = "" Then gSettings.LabelI1I2 = "I1/I2"
'
'
'
'
'    Call Utility_Search("RUNIDLOOKUPTYPE", "", "NAME", True, True, False)
'    gSettings.RunIDLookupType = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("SEEDPREFIXCT", "", "NAME", True, False, False)
'    gSettings.CTPrefix = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("SEEDVALUECT", "", "NAME", True, True, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.CTSeedValue = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.CTSeedValue = 1
'    End If
'
'    Call Utility_Search("SEEDPREFIXLR", "", "NAME", True, True, False)
'    gSettings.LRPrefix = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("SEEDVALUELR", "1000", "NAME", True, True, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.LRSEEDValue = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.LRSEEDValue = 1000
'    End If
'
'
'    Call Utility_Search("SEEDVALUETAG-LR", "1000", "NAME", True, True, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.LRSeedValueTag = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.LRSeedValueTag = 250000
'    End If
'
'    Call Utility_Search("SEEDPREFIXTAG", "", "NAME", True, False, False)
'    gSettings.SEEDPREFIXTAG = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("SEEDVALUETAG-K", "", "NAME", True, False, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.KDSeedVAlueTag = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.KDSeedVAlueTag = 800000
'    End If
'
'    Call Utility_Search("SEEDVALUETAG-G", "", "NAME", True, False, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.GRSeedValueTag = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.GRSeedValueTag = 400000
'    End If
'
'    If gSettings.GRSeedValueTag < 100 Then gSettings.GRSeedValueTag = 101
'
'
'        Call Utility_Search("SEEDVALUETAG-K", "", "NAME", True, False, False)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gSettings.ADSeedValueTag = CLng(gudtUtility(0).UtilityValueText)
'    Else
'        gSettings.ADSeedValueTag = 600000
'    End If
'
'
'    Call Utility_Search("BTCOMMRECEIVE", "", "NAME", True, True)
'    glngBTCommReceive = gudtUtility(0).UtilityValueLong
'
'    Call Utility_Search("RECEIVEESTIMATEMETHOD", "", "NAME", True, True)
'    gstrReceiveEstimateMethod = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("SHIPOVERAGEPERCENT", "", "NAME", True, True)
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gdblShipOveragePercent = CDbl(gudtUtility(0).UtilityValueText)
'    Else
'        gdblShipOveragePercent = 0
'    End If
'
'    Call Utility_Search("ENDTALLYMAXWIDTH", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gintEndTallyMaxWidth = CInt(gudtUtility(0).UtilityValueText)
'    Else
'        gintEndTallyMaxWidth = 29
'    End If
'
'    Call Utility_Search("ENDTALLYTWOKEYMAX", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
'        gintEndTallyTwoKeyMax = CInt(gudtUtility(0).UtilityValueText)
'    Else
'        gintEndTallyTwoKeyMax = 2
'    End If
'
'    Call Utility_Search("ETKEYF3LOCATION", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.ETKeyF3Location = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("ETKEYF4LOCATION", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.ETKeyF4Location = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("ETKEYF5LOCATION", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.ETKeyF5Location = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("ETKEYF6LOCATION", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.ETKeyF6Location = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("RECEIVEDETAILAUTOPRINT", "0", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.ReceiveDetailAutoPrint = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("LOGINMETHOD", "0", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.LoginMethod = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("RECEIVEDAYLOADOPTION", "5", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.ReceiveDayLoadOption = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("RECEIVEDETAILSCANTAG", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.ReceiveDetailScanTag = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("DEFAULTTHICKNESS", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.DefaultThickness = gudtUtility(0).UtilityValueText
'    gstrLastThickness = gSettings.DefaultThickness
'
'    Call Utility_Search("MACAID", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.MACAID = gudtUtility(0).UtilityValueText
'
'    Call Utility_Search("BLOCKTALLYCUSTOM", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.BlockTallyCustom = Trim(UCase(gudtUtility(0).UtilityValueText))
'
'
'
'    Call Utility_Search("LASTRUNID", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.LastRUNID = Trim(UCase(gudtUtility(0).UtilityValueText))
'    gstrLastProdRunAID = gSettings.LastRUNID
'
'    Call Utility_Search("DEFAULTDIMENSIONED", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'
'    If IsNumeric(Trim(UCase(gudtUtility(0).UtilityValueText))) = True Then
'        gSettings.DefaultDimensioned = CInt(Trim(UCase(gudtUtility(0).UtilityValueText)))
'    Else
'        gSettings.DefaultDimensioned = 0
'    End If
'    gstrLastProdRunAID = gSettings.LastRUNID
'
'
'    Call Utility_Search("BUNDLETODAYESTFORM", "", "NAME", True, True, False, "UTILITYVALUETEXT")
'    gSettings.BundleTodayESTForm = Trim(UCase(gudtUtility(0).UtilityValueText))
'
'
'
'    Exit Sub
''
'ErrorHandler:
'    MsgBox "Error during setting initialization, Contact Technical Support at 888.520.1951 (" & Err.Number & "-" & Err.Description & ")"
'    Exit Sub
''
'End Sub
Public Function AV(strString As String) As Double
    If IsNumeric(strString) = True Then
        AV = CDbl(strString)
    Else
        AV = 0
    End If
End Function
Public Function CLngAV(strString As String) As Long
    CLngAV = CLng(AV(strString))
End Function
Public Function GetHHSerialAdjustment(strSerial As String) As String

    
On Error GoTo ErrorHandler
    KeyHelp2_99EX = 112
    KeyDelete = 189
    
    If Trim(UCase(strSerial)) = "006E604F30142" Or SC(strSerial, "006E604F30141") = True Or Trim(UCase(strSerial)) = "006E604F30134" Or _
        Trim(UCase(strSerial)) = "006E604F30138" Or Trim(UCase(strSerial)) = "006E604F30143" Or _
        Trim(UCase(strSerial)) = "006E604F30131" Or Trim(UCase(strSerial)) = "50F0063006B000000" Then   'its a duplicated one
        
        
        If tcu(strSerial) = "50F0063006B000000" Then
            gf9900 = False
            KeyDelete = 46
            If SC(gSettings.KeySpaceAsF1, "YES") Then
                KeyHelp2_99EX = 32
            ElseIf IsNumeric(gSettings.KeySpaceAsF1) = True Then
                KeyHelp2_99EX = CInt(gSettings.KeySpaceAsF1)
            End If
        End If
        
        Call Utility_Search("HWSERIAL", "", "NAME", True, False, False, "")
        
        strSerial = strSerial & Trim(UCase(gudtUtility(0).UtilityValue2Text))
    End If
    
    GetHHSerialAdjustment = strSerial
    
    Exit Function
ErrorHandler:
    MsgBox "Could not access Utility.pdb for Serial#"
    Exit Function
    
End Function

Public Sub KeyBeep(Keytone As AFTone)
       
  Keytone.Pitch = 300
  Keytone.Duration = 100
  Keytone.Play
  
End Sub
Public Sub KeyBeepError(Keytone As AFTone, Optional strType As String, Optional fSpecialSpeed As Boolean)
    Dim lng10 As Long, lng9 As Long, lng50 As Long, lng40 As Long, lng150 As Long
    
'    If tcu(App.Path) Like "\\VM*" Or tcu(App.Path) Like "C:\APP*" Then
'        gSettings.ETKeySound = "fast"
'    End If
    
    If fSpecialSpeed = True Or SC(gSettings.ETKeySound, "fast") = True Or SC(gSettings.ETKeySound, "FAST2") = True Then
        If fSpecialSpeed = True Then
            If CLng(gSettings.SoundSpeed9) <> 0 Then lng9 = CLng(gSettings.SoundSpeed9) * 4
            If CLng(gSettings.SoundSpeed10) <> 0 Then lng10 = CLng(gSettings.SoundSpeed10) * 5
            If CLng(gSettings.SoundSpeed40) <> 0 Then lng40 = CLng(gSettings.SoundSpeed40) * 20
            If CLng(gSettings.SoundSpeed50) <> 0 Then lng50 = CLng(gSettings.SoundSpeed50) * 30
            If CLng(gSettings.SoundSpeed150) <> 0 Then lng150 = CLng(gSettings.SoundSpeed150) * 90
            
        ElseIf SC(gSettings.ETKeySound, "FAST") = True Then
            lng9 = 9
            lng10 = 10
            lng40 = 40
            lng50 = 50
            lng150 = 150
        ElseIf SC(gSettings.ETKeySound, "FAST2") = True Then
            lng9 = 4
            lng10 = 5
            lng40 = 20
            lng50 = 30
            lng150 = 90
        End If
        
        If strType = "" Then
          Keytone.Pitch = 100
          Keytone.Duration = lng10
          Keytone.Play
        
          Keytone.Pitch = 500
          Keytone.Duration = lng9
          Keytone.Play
          Keytone.Pitch = 2000
          Keytone.Duration = lng10
          Keytone.Play
          
          Keytone.Pitch = 3000
          Keytone.Duration = lng10
          Keytone.Play
          
        ElseIf strType = "MAJORERROR" Then
          Keytone.Pitch = 1200
          Keytone.Duration = lng50
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = lng10
          Keytone.Play
          
          Keytone.Pitch = 500
          Keytone.Duration = lng10
          Keytone.Play
        
          Keytone.Pitch = 3200
          Keytone.Duration = lng50
          Keytone.Play
        
          Keytone.Pitch = 5200
          Keytone.Duration = lng50
          Keytone.Play
          
          Keytone.Pitch = 1200
          Keytone.Duration = lng50
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = lng40
          Keytone.Play
          
          Keytone.Pitch = 500
          Keytone.Duration = lng40
          Keytone.Play
        
          Keytone.Pitch = 3200
          Keytone.Duration = lng150
          Keytone.Play
    '      Keytone.Pitch = 3000
    '      Keytone.Duration = 55
    '      Keytone.Play
    '
    '      Keytone.Pitch = 6200
    '      Keytone.Duration = 75
    '      Keytone.Play
    '
    '      Keytone.Pitch = 200
    '      Keytone.Duration = 75
    '      Keytone.Play
        ElseIf strType = "SUCCESS" Then
          Keytone.Pitch = 300
          Keytone.Duration = lng9
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = lng9
          Keytone.Play
          
          Keytone.Pitch = 2200
          Keytone.Duration = lng9
          Keytone.Play
          
          Keytone.Pitch = 3200
          Keytone.Duration = lng9
          Keytone.Play
          
          Keytone.Pitch = 5200
          Keytone.Duration = lng9
          Keytone.Play
          
          Keytone.Pitch = 6200
          Keytone.Duration = lng9
          Keytone.Play
        
        ElseIf strType = "ERROR" Then
          Keytone.Pitch = 5200
          Keytone.Duration = lng9
          Keytone.Play
        
          Keytone.Pitch = 2200
          Keytone.Duration = lng9
          Keytone.Play
          
          Keytone.Pitch = 3200
          Keytone.Duration = lng9
          Keytone.Play
          
          Keytone.Pitch = 1000
          Keytone.Duration = lng40
          Keytone.Play
          
          Keytone.Pitch = 300
          Keytone.Duration = lng40
          Keytone.Play
          
        End If
    ElseIf gf99EX = False And SC(gstrModel, "99EX") = False Then
        If strType = "" Then
          Keytone.Pitch = 300
          Keytone.Duration = 30
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 20
          Keytone.Play
          Keytone.Pitch = 5200
          Keytone.Duration = 30
          Keytone.Play
          
          Keytone.Pitch = 6200
          Keytone.Duration = 30
          Keytone.Play
          
        ElseIf strType = "MAJORERROR" Then
          Keytone.Pitch = 1200
          Keytone.Duration = 50
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 30
          Keytone.Play
          
          Keytone.Pitch = 500
          Keytone.Duration = 30
          Keytone.Play
        
          Keytone.Pitch = 3200
          Keytone.Duration = 50
          Keytone.Play
        
          Keytone.Pitch = 5200
          Keytone.Duration = 50
          Keytone.Play
          
          Keytone.Pitch = 1200
          Keytone.Duration = 50
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 30
          Keytone.Play
          
          Keytone.Pitch = 500
          Keytone.Duration = 30
          Keytone.Play
        
          Keytone.Pitch = 3200
          Keytone.Duration = 30
          Keytone.Play
    '      Keytone.Pitch = 3000
    '      Keytone.Duration = 55
    '      Keytone.Play
    '
    '      Keytone.Pitch = 6200
    '      Keytone.Duration = 75
    '      Keytone.Play
    '
    '      Keytone.Pitch = 200
    '      Keytone.Duration = 75
    '      Keytone.Play
        ElseIf strType = "SUCCESS" Then
          Keytone.Pitch = 300
          Keytone.Duration = 20
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 2200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 3200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 5200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 6200
          Keytone.Duration = 20
          Keytone.Play
        
        ElseIf strType = "ERROR" Then
          Keytone.Pitch = 5200
          Keytone.Duration = 20
          Keytone.Play
        
          Keytone.Pitch = 2200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 3200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 1000
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 300
          Keytone.Duration = 20
          Keytone.Play
          
        End If
    
    ElseIf gf99EX = True Or SC(gstrModel, "99EX") = True Then
        If strType = "" Then
          Keytone.Pitch = 300
          Keytone.Duration = 30
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 20
          Keytone.Play
          Keytone.Pitch = 5200
          Keytone.Duration = 30
          Keytone.Play
          
          Keytone.Pitch = 6200
          Keytone.Duration = 30
          Keytone.Play
          
        ElseIf strType = "MAJORERROR" Then
          Keytone.Pitch = 1200
          Keytone.Duration = 50
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 30
          Keytone.Play
          
          Keytone.Pitch = 500
          Keytone.Duration = 30
          Keytone.Play
        
          Keytone.Pitch = 3200
          Keytone.Duration = 50
          Keytone.Play
        
          Keytone.Pitch = 5200
          Keytone.Duration = 50
          Keytone.Play
          
          Keytone.Pitch = 1200
          Keytone.Duration = 50
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 300
          Keytone.Play
          
          Keytone.Pitch = 500
          Keytone.Duration = 300
          Keytone.Play
        
          Keytone.Pitch = 3200
          Keytone.Duration = 500
          Keytone.Play
    '      Keytone.Pitch = 3000
    '      Keytone.Duration = 55
    '      Keytone.Play
    '
    '      Keytone.Pitch = 6200
    '      Keytone.Duration = 75
    '      Keytone.Play
    '
    '      Keytone.Pitch = 200
    '      Keytone.Duration = 75
    '      Keytone.Play
        ElseIf strType = "SUCCESS" Then
          Keytone.Pitch = 300
          Keytone.Duration = 20
          Keytone.Play
        
          Keytone.Pitch = 1000
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 2200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 3200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 5200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 6200
          Keytone.Duration = 20
          Keytone.Play
        
        ElseIf strType = "ERROR" Then
          Keytone.Pitch = 5200
          Keytone.Duration = 20
          Keytone.Play
        
          Keytone.Pitch = 2200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 3200
          Keytone.Duration = 20
          Keytone.Play
          
          Keytone.Pitch = 1000
          Keytone.Duration = 200
          Keytone.Play
          
          Keytone.Pitch = 300
          Keytone.Duration = 200
          Keytone.Play
          
        End If
        
    End If
    
End Sub
Public Sub GetPrinterSettings()
    gstrPrinterCommMethod = "BLUETOOTH"
    gstrPrinterIP = "192.168.25.200"
    
    Call Utility_Search("BTCOMM", "", "NAME", True, False)
    
    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
        gintPrinterPort = CLng(gudtUtility(0).UtilityValueText)
    Else
        gintPrinterPort = 7
    End If
    

End Sub
Public Sub UpdateSpeciesList(cbo As AFComboBox, Optional fArrayOnly As Boolean)
    Dim strText As String
    Dim dblWeight As Double
    Dim lngID As Long
    Dim strHHAID As String
    Dim I As Integer
    Dim udt As tSpeciesRecord
    
   OpenSpeciesDatabase     ' ltas 8.29.2006
   
    If dbSpecies = 0 Then
        MsgBox "Unable to open Species database"
        Exit Sub
    End If
    
    PDBSetSortFields dbSpecies, 2
    
    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If
    
    If fFirstSpeciesLoad = True Then
        I = -1
    Else
        ReDim garySpecies(3, 0)
        I = 0
        fFirstSpeciesLoad = True
    End If
    
    PDBMoveFirst dbSpecies
    While Not PDBEOF(dbSpecies)
        PDBReadRecord dbSpecies, VarPtr(udt)
        
        
        If fArrayOnly = False Then
            cbo.AddItem udt.SpeciesDescription
            cbo.ItemData(cbo.NewIndex) = udt.SpeciesID
        End If
        
        If I <> -1 Then
            If I = 0 And garySpecies(0, 0) = "" Then
                'Do Nothing..noredim needed
            Else
                I = I + 1
                ReDim Preserve garySpecies(3, I)
            End If
            garySpecies(0, I) = udt.SpeciesID
            garySpecies(1, I) = udt.SpeciesDescription
            garySpecies(2, I) = udt.SpeciesAbbrev
            garySpecies(3, I) = udt.SpeciesWeight
        End If
        
        PDBMoveNext dbSpecies
    Wend
    PDBClose dbSpecies
        
        
    If fArrayOnly = False Then cbo.ListIndex = 0
    
End Sub
Public Function GetSpeciesData(strSearchString As String, strSearchType As String, Optional strReturnField As String, _
                                Optional ByRef strReturnSpeciesHHAID As String, _
                                Optional ByRef strReturnSpeciesName As String, _
                                Optional lngReturnID As Long) As String
    Dim I As Integer
    Dim fFound As Boolean
        
On Error GoTo ErrorHandler

    If strReturnField = "" Then
        GetSpeciesData = "-1,INVALID, ,0"
    Else
        GetSpeciesData = "INVALID"
    End If
    strReturnSpeciesHHAID = "-1"
    strReturnSpeciesName = "INVALID"
    
    If SC(strSearchString, "") = True Then
        'don't bother checking, it's blank
        'it's blank so it's not found..
        fFound = False

    Else
        For I = 0 To UBound(garySpecies, 2)
            If SC(strSearchType, "AID") = True Or SC(strSearchType, "HHAID") = True Or SC(strSearchType, "ABBREV") = True Or SC(strSearchType, "SPECIESABBREV") = True Then
                If Trim(UCase(garySpecies(2, I))) = Trim(UCase(strSearchString)) Then
                    'Match Found
                    fFound = True
                End If
            ElseIf SC(strSearchType, "SPECIESNAME") Or SC(strSearchType, "SPECIESDESCRIPTION") Or SC(strSearchType, "SPECIES") Or SC(strSearchType, "NAME") = True Or SC(strSearchType, "DESC") = True Then
                If Trim(UCase(garySpecies(1, I))) = Trim(UCase(strSearchString)) Then
                    'Match Found
                    fFound = True
                End If
            Else
                If Trim(UCase(garySpecies(0, I))) = Trim(UCase(strSearchString)) Then
                    'Match Found
                    fFound = True
                End If
            End If
            
            If fFound = True Then
                strReturnSpeciesHHAID = garySpecies(2, I)
                strReturnSpeciesName = garySpecies(1, I)
                If IsNumeric(garySpecies(0, I)) = True Then lngReturnID = CLng(garySpecies(0, I))
                
                If strReturnField = "" Then
                    GetSpeciesData = garySpecies(0, I) & "," & garySpecies(1, I) & "," & garySpecies(2, I) & "," & garySpecies(3, I)
                Else
                    If SC(strReturnField, "HHAID") Or SC(strReturnField, "AID") Then
                        GetSpeciesData = garySpecies(2, I)
                    ElseIf SC(strReturnField, "ID") Then
                        GetSpeciesData = garySpecies(0, I)
                    ElseIf SC(strReturnField, "SPECIES") Or SC(strSearchType, "SPECIESNAME") Or SC(strReturnField, "SPECIESDESCRIPTION") Or SC(strReturnField, "NAME") Or SC(strReturnField, "DESC") Then
                        GetSpeciesData = garySpecies(1, I)
                    ElseIf SC(strReturnField, "LBS") = True Then
                        GetSpeciesData = garySpecies(3, I)
                    End If
                End If
                Exit For
            End If
                        
        Next
    End If
    
    If fFound = False Then
        strReturnSpeciesHHAID = "NA"
        strReturnSpeciesName = "NOT FOUND"
        lngReturnID = -1
        
        If strReturnField = "ID" Then
            GetSpeciesData = "-1"
        ElseIf (SC(strReturnField, "HHAID") Or SC(strReturnField, "AID")) Then
            GetSpeciesData = "-1"
        Else
            'if return field is blank then send back the comma delimited string
            GetSpeciesData = "-1,INVALID, ,0"
        End If
    End If
    
    Exit Function
ErrorHandler:
    GetSpeciesData = "-1"
    MsgBox "Error in GetSpeciesData " & Err.Number & "-" & Err.Description
    Exit Function
End Function
Public Sub UpdateBEList(cbo As AFComboBox, Optional fArrayOnly As Boolean)
    
    Dim udt As tBERecord
    Dim I As Integer
    
    ReDim garyBE(0)
    
   OpenBEDatabase      ' ltas 8.29.2006
    If dbBE = 0 Then
        MsgBox "Unable to open BE database"
        Exit Sub
    End If
    
    PDBSetSortFields dbBE, 0
    
    If fFirstBELoad = True Then
        I = -1
    Else
        ReDim garyBE(0)
        I = 0
        fFirstBELoad = True
    End If
    
    PDBMoveFirst dbBE
    While Not PDBEOF(dbBE)
        
        I = I + 1
        
        If I = -1 Then
            ReDim garyBE(0)
            garyBE(0).EstimatedBundleFootageID = -1
        Else
            ReDim Preserve garyBE(I)
        End If
        
        ReadBERecord dbBE, garyBE(I)
        
        If PDBGetLastError(dbBE) = ErrNone Then ' continue
        
        Else
            MsgBox "ERROR READING BE DATABASE!"
        End If
        
        
        If I <> -1 Then
            If I = 0 And garyBE(0).EstimatedBundleFootageID = -1 Then
                'Do Nothing..noredim needed
            Else
                I = I + 1
                ReDim Preserve garyBE(I)
            End If
            garyBE(I).EstimatedBundleFootageID = udt.EstimatedBundleFootageID
            garyBE(I).FootagePerLayer = udt.FootagePerLayer
            garyBE(I).SPEC_GR_THK_LN = udt.SPEC_GR_THK_LN
        End If
        
        PDBMoveNext dbBE
    Wend
    
    PDBClose dbBE
    dbBE = 0
    
    If fArrayOnly = False Then cbo.ListIndex = 0
    
End Sub
Public Sub UpdateMACList(cbo As AFComboBox, Optional fArrayOnly As Boolean)
    
    Dim udt As tMACRecord
    Dim I As Integer
    
    ReDim garyMAC(0)
    garyMAC(0).MACID = 0
    
    OpenMACDatabase
    
    If dbMAC = 0 Then
        MsgBox "Unable to open MAC database"
        PDBClose dbMAC
        dbMAC = 0
        Exit Sub
    End If
    
    
    PDBSetSortFields dbMAC, 1
    PDBMoveFirst dbMAC
    
    I = -1
    While Not PDBEOF(dbMAC)
        I = I + 1
        
        If I = 0 Then
            ReDim garyMAC(0)
            garyMAC(0).MACID = 0
        Else
            ReDim Preserve garyMAC(I)
        End If
        
        ReadMACRecord garyMAC(I)
        Debug.Print garyMAC(I).MACAID & "-" & garyMAC(I).MACDesc
        
        If PDBGetLastError(dbMAC) = ErrNone Then ' continue
        
        Else
            MsgBox "ERROR READING MAC DATABASE!"
        End If
            
        PDBMoveNext dbMAC
    Wend
    
    PDBClose dbMAC
    dbMAC = 0
    
    If fArrayOnly = False Then cbo.ListIndex = 0
    
End Sub
Public Sub UpdateProdRunList()
    
    Dim I As Integer
    Dim lngRecCount As Long
    
    ReDim garyProdRun(0)
    garyProdRun(0).ProdRunID = 0
    
    OpenProdRunDatabase
    
    If dbProdRun = 0 Then
        MsgBox "Unable to open ProdRun database"
        PDBClose dbProdRun
        dbProdRun = 0
        Exit Sub
    End If
    
    
    PDBSetSortFields dbProdRun, 1
    PDBMoveFirst dbProdRun
    
    If PDBNumRecords(dbProdRun) > 0 Then
        ReDim garyProdRun(PDBNumRecords(dbProdRun) - 1)
    Else
        ReDim garyProdRun(0)
        garyProdRun(0).ProdRunAID = ""
    End If
    
    lngRecCount = PDBBulkRead(dbProdRun, PDBNumRecords(dbProdRun), VarPtr(garyProdRun(0)))
    
    PDBClose dbProdRun
    dbProdRun = 0
    
    
    
End Sub

Public Sub UpdateLOCList(cbo As AFComboBox, Optional fArrayOnly As Boolean)
    
    Dim udt As tLOCRecord
    Dim I As Integer
   
On Error GoTo ErrorHandler

    OpenLOCDatabase
    ReDim garyLoc(0)
    garyLoc(0).LOCID = 0
    
    If dbLOC = 0 Then
        MsgBox "Unable to open LOC database"
        Exit Sub
    End If
    
    PDBSetSortFields dbLOC, 1
    
    I = -1
    
    PDBMoveFirst dbLOC
    While Not PDBEOF(dbLOC)
        
        
        
        I = I + 1
        
        If I = 0 Then
            ReDim garyLoc(0)
            garyLoc(0).LOCID = 0
        Else
            ReDim Preserve garyLoc(I)
        End If
        
        ReadLOCRecord garyLoc(I)
        Debug.Print garyLoc(I).LocAID & "-" & garyLoc(I).LOCDesc
        
        If PDBGetLastError(dbLOC) = ErrNone Then ' continue
        
        Else
            MsgBox "ERROR READING LOC DATABASE!"
        End If
                
        PDBMoveNext dbLOC
    Wend
    
    PDBClose dbLOC
    dbLOC = 0
    
    If fArrayOnly = False Then cbo.ListIndex = 0
    Exit Sub
ErrorHandler:
    MsgBox "Error Loading the Location List to mLIMBS Memory/Lookup Array Error# " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub

Public Sub UpdateUserList(cbo As AFComboBox, Optional fArrayOnly As Boolean)
    
    Dim strUserName As String, strPassword As String
    Dim strSecurity As String
    Dim lngID As Long
    Dim I As Integer
    Dim udt As tUserRecord
    
  OpenUserDatabase     ' ltas 8.29.2006
    If dbUser = 0 Then
        MsgBox "Unable to open User database"
        Exit Sub
    End If
    
    PDBSetSortFields dbUser, 1
    
    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If
    
    If fFirstUserLoad = True Then
        I = -1
    Else
        ReDim garyUser(3, 0)
        I = 0
        fFirstUserLoad = True
    End If
    
    PDBMoveFirst dbUser
    While Not PDBEOF(dbUser)
        PDBReadRecord dbUser, VarPtr(udt)
        
        
        If Trim(UCase(udt.UserLoginNameHH)) = "" Then
            'do nothing
        Else
            If fArrayOnly = False Then
                cbo.AddItem udt.UserLoginNameHH
                cbo.ItemData(cbo.NewIndex) = udt.UserID
            End If
            
            If I <> -1 Then
                If I = 0 And garyUser(0, 0) = "" Then
                    'Do Nothing..noredim needed
                Else
                    I = I + 1
                    ReDim Preserve garyUser(3, I)
                End If
                garyUser(0, I) = udt.UserID
                garyUser(1, I) = udt.UserLoginNameHH
                garyUser(2, I) = udt.UserPasswordHH
                garyUser(3, I) = udt.UserSecurityHH
            End If
        End If
        PDBMoveNext dbUser
    Wend
    PDBClose dbUser
        
    If fArrayOnly = False Then cbo.ListIndex = 0
    
End Sub
Public Function GetUserData(strSearchString As String, strSearchType As String) As String
    Dim I As Integer
    Dim fFound As Boolean
    
    GetUserData = "-1,INVALID, ,0"
    For I = 0 To UBound(garyUser, 2)
        If strSearchType = "USERNAME" Then
            If Trim(UCase(garyUser(1, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        Else
            If Trim(UCase(garyUser(0, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        End If
        
        If fFound = True Then
            GetUserData = garyUser(0, I) & "," & garyUser(1, I) & "," & garyUser(2, I) & "," & garyUser(3, I)
            Exit For
        End If
                    
    Next
    
    Exit Function
       
End Function

Public Sub UpdatePAList(cbo As AFComboBox, Optional fArrayOnly As Boolean, Optional fSingleItem As Boolean, Optional lngPAID As Long)
    Dim db As Long
    Dim strText As String
    Dim strText2 As String
    Dim strText3 As String
    Dim lngID As Long
    Dim I As Integer
    

    db = PDBOpen(Byfilename, gstrPDBPath & "\PA", 0, 0, 0, 0, afModeReadWrite)

    If db = 0 Then
        MsgBox "Unable to open PA database"
        Exit Sub
    End If

    If fSingleItem = True Then
        PDBSetSortFields db, tPADatabaseFields.PAID_Field
    Else
        PDBSetSort db, "PAType,DisplayOrder,PAName"
    End If
    
    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If

    If fFirstPALoad = True Then
        I = -1
        ReDim garyPA(0)
    Else
        ReDim garyPA(0)
        I = -1
        fFirstPALoad = True
    End If

    PDBMoveFirst db
    
    If fSingleItem = True Then
        PDBFindRecordByField db, tPADatabaseFields.PAID_Field, lngPAID
        If PDBGetLastError(db) <> 0 Then
            MsgBox "PA Record for ID: " & lngPAID & " Could Not be Found"
            Exit Sub
        End If
    End If
    
    
    While Not PDBEOF(db)
        I = I + 1
        If I = 0 Then
            ReDim garyPA(0)
        Else
            ReDim Preserve garyPA(I)
        End If
        
        PDBReadRecord db, VarPtr(garyPA(I))
        
        If fArrayOnly = False Then
            cbo.AddItem garyPA(I).PAName
            cbo.ItemData(cbo.NewIndex) = garyPA(I).PAID
        End If

        If fSingleItem = True Then PDBMoveLast db
        PDBMoveNext db
    Wend
    PDBClose db

    If fArrayOnly = False Then cbo.ListIndex = 0

End Sub

Public Function GetPAData(strSearchString As String, strSearchType As String, Optional strPAType As String, _
    Optional fRequireTypeMatch As Boolean, Optional strReturnSingle As String) As String
    Dim I As Integer
    Dim fFound As Boolean
 
On Error GoTo ErrorHandler

    GetPAData = "-1,INVALID,"
    For I = 0 To UBound(garyPA)
        If ((strSearchType = "PANAME" And SC(strSearchString, garyPA(I).PAName) = True) Or (strSearchType = "HHAID" And SC(strSearchString, garyPA(I).PAName) = True)) And _
            (fRequireTypeMatch = False Or (fRequireTypeMatch = True And Trim(UCase(strPAType)) = Trim(UCase(garyPA(I).PAType)))) Then
                'Match Found
            fFound = True
        
        ElseIf IsNumeric(strSearchString) And (fRequireTypeMatch = False Or (tcu(strSearchString) = tcu(garyPA(I).PAName) And fRequireTypeMatch = True And Trim(UCase(strPAType)) = Trim(UCase(garyPA(I).PAType)))) Then
            If garyPA(I).PAID = CLng(strSearchString) Then
                'Match Found
                fFound = True
            End If
        End If
        
        If fFound = True Then
            If strReturnSingle = "HHAID" Then
                GetPAData = garyPA(I).PAName
            ElseIf SC(strReturnSingle, "PADESC") = True Then
                GetPAData = garyPA(I).PADesc
            ElseIf SC(strReturnSingle, "ID") = True Then
                GetPAData = CStr(garyPA(I).PAID)
            Else
                GetPAData = garyPA(I).PAID & "," & garyPA(I).PAName & "," & garyPA(I).PADesc
            End If
            Exit For
        End If
                    
    Next
    
    
    If (SC(strReturnSingle, "HHAID") Or SC(strReturnSingle, "PANAME")) And fFound = False Then
        GetPAData = "-1"
    ElseIf SC(strReturnSingle, "PADESC") And fFound = False Then
        GetPAData = strSearchString
    ElseIf strReturnSingle <> "" And fFound = False Then
        GetPAData = "-1"
    End If
    
    Exit Function
ErrorHandler:
    GetPAData = "-1"
    
    MsgBox "Error in GetPAData Search String: " & strSearchString & " Search type=" & strSearchType & " PAType=" & strPAType & vbCrLf & "Error #" & Err.Number & " - " & Err.Description
    
    
End Function
Public Function SC(strString1 As String, strString2 As String) As Boolean
    If Trim(UCase(strString1)) = Trim(UCase(strString2)) Then
        SC = True
    Else
        SC = False
    End If
End Function
Public Function VC(strString1 As String, strString2 As String) As Boolean
    Dim dblString1 As Double, dblString2 As Double
    dblString1 = 0
    dblString2 = 0
    
    If IsNumeric(Trim(strString1)) = False Then
        dblString1 = 0
    Else
        dblString1 = CDbl(Trim(strString1))
    End If
    
    If IsNumeric(Trim(strString2)) = False Then
        dblString2 = 0
    Else
        dblString2 = CDbl(Trim(strString2))
    End If
        
    If dblString1 = dblString2 Then
        VC = True
    Else
        VC = False
    End If
    
    Exit Function
    
End Function
Public Function ValA(strString1 As String) As Double
    Dim dblString1 As Double

On Error GoTo ErrorHandler

    dblString1 = 0
    
    If IsNumeric(Trim(strString1)) = False Then
        dblString1 = 0
    Else
        dblString1 = CDbl(Trim(strString1))
    End If
    
    ValA = dblString1
    
    Exit Function
ErrorHandler:
    MsgBox "Error in ValA(string1=" & strString1 & ") Error# " & Err.Number & "-" & Err.Description
    Exit Function
    
End Function
Public Function SCInList(str1 As String, strValList As String, Optional strDelimeter As String) As Boolean
    
    Dim I As Integer
    Dim strSPLIT_Local As aryStringType
    
    If SC(strDelimeter, "") = True Then strDelimeter = ","
    
    Call AppForge_Split(strValList, strSPLIT_Local, strDelimeter)
    
    
    If UBound(strSPLIT_Local.ary) = 0 And strSPLIT_Local.ary(0) = "" Then
        strSPLIT_Local.ary(0) = strValList
    End If
    
    For I = 0 To UBound(strSPLIT_Local.ary)
        If SC(str1, CStr(strSPLIT_Local.ary(I))) = True Then
            SCInList = True
            Exit Function
        End If
    Next
    
    Exit Function
ErrorHandler:
    SCInList = False
    
    MsgBox "Error in SCInList(" & str1 & " in? " & strValList & " Separated by " & strDelimeter & " - " & Err.Number & " " & Err.Description
    Exit Function
    
End Function

Public Function tc(strString As String) As String
    tc = (Trim(strString))
End Function

Public Function tcu(ByVal strString As String) As String
    tcu = UCase(Trim(strString))
End Function
Public Sub UpdateThicknessList(cbo As AFComboBox, Optional fArrayOnly As Boolean)

On Error GoTo ErrorHandler

    Dim strText As String
    Dim lngID As Long
    Dim strHHAID As String
    Dim I As Integer
    Dim udt As tThicknessRecord
    
    OpenThicknessDatabase     ' ltas 8.29.2006
    
    If dbThickness = 0 Then
        MsgBox "Unable to open Thickness database"
        Exit Sub
    End If
    
    PDBSetSortFields dbThickness, 1
    
    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If
    
    If fFirstThicknessLoad = True Then
        I = -1
    Else
        ReDim garyThickness(3, 0)
        I = 0
        fFirstThicknessLoad = True
    End If
    
    PDBMoveFirst dbThickness
    
    While Not PDBEOF(dbThickness)
        PDBReadRecord dbThickness, VarPtr(udt)
                
        If fArrayOnly = False Then
            cbo.AddItem udt.thickness
            cbo.ItemData(cbo.NewIndex) = udt.ThicknessID
        End If
        
        If I <> -1 Then
            If I = 0 And garyThickness(0, 0) = "" Then
                'Do Nothing..noredim needed
            Else
                I = I + 1
                ReDim Preserve garyThickness(3, I)
            End If
            garyThickness(0, I) = udt.ThicknessID
            garyThickness(1, I) = udt.thickness
            garyThickness(2, I) = udt.HHAID
            garyThickness(3, I) = udt.CalcFactor
        End If
        
        PDBMoveNext dbThickness
    Wend
    
    PDBClose dbThickness
        
    If fArrayOnly = False Then cbo.ListIndex = 0
    Exit Sub
ErrorHandler:
    MsgBox "Error in UpdateThicknessList: " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub
Public Function GetThicknessData(strSearchString As String, strSearchType As String, Optional strReturnField As String, _
                                Optional lngReturnID As Long, Optional strReturnAID As String, _
                                Optional strReturnName As String) As String
    Dim I As Integer
    Dim fFound As Boolean
On Error GoTo ErrorHandler

    GetThicknessData = "-1,INVALID,0,0"
    
    For I = 0 To UBound(garyThickness, 2)
        If SC(strSearchType, "AID") Or SC(strSearchType, "HHAID") Then
            If Trim(UCase(garyThickness(2, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "Thickness") = True Or SC(strSearchType, "THK") = True Or SC(strSearchType, "NAME") = True Or SC(strSearchType, "DESC") = True Then
            If Trim(UCase(garyThickness(1, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "ID") Then
            If Trim(UCase(garyThickness(0, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        
        Else
            If Trim(UCase(garyThickness(0, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        End If
        
        If fFound = True Then
            lngReturnID = CLng(garyThickness(0, I))
            strReturnAID = garyThickness(2, I)
            strReturnName = garyThickness(1, I)
            
                
            If strReturnField = "" Then
                GetThicknessData = garyThickness(0, I) & "," & garyThickness(1, I) & "," & garyThickness(2, I) & "," & garyThickness(3, I)
                
            ElseIf SC(strReturnField, "AID") Or SC(strReturnField, "HHAID") Then
                GetThicknessData = garyThickness(2, I)
                
            ElseIf Trim(UCase(strReturnField)) = "CALCFACTOR" Then
                If IsNumeric(garyThickness(3, I)) Then
                    GetThicknessData = CStr(CDbl(garyThickness(3, I)))
                Else
                    GetThicknessData = "0"
                End If
                
            ElseIf strReturnField = "ID" Then
                GetThicknessData = (garyThickness(0, I))
                
            ElseIf SC(strReturnField, "Thickness") Or SC(strReturnField, "THK") Or SC(strReturnField, "NAME") Or SC(strReturnField, "DESC") Then
                GetThicknessData = garyThickness(1, I)
            End If
            'ffound is true so exit for
            Exit For
        End If
    Next
    
    If fFound = False Then
        strReturnAID = "NA"
        strReturnName = "NOT FOUND"
        lngReturnID = -1
    
        If Trim(UCase(strReturnField)) = "CALCFACTOR" Then
            GetThicknessData = 0
        ElseIf Trim(UCase(strReturnField)) <> "" Then
            GetThicknessData = -1
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error in GetThicknessData " & Err.Number & "-" & Err.Description
    Exit Function
    
End Function

Public Sub UpdateGradeList(cbo As AFComboBox, Optional fArrayOnly As Boolean)
    Dim I As Integer
    
    Dim strGrade As String
    Dim lngGradeID As Long
    Dim strHHAID As String
    Dim intGradingUse As Integer
    Dim intDisplayOrder As Integer
    Dim udt As tGradeRecord
    
    'Write event log for trouble shooting
'    '     Call WriteEventLog(0, "UPDATEGRADELIST", "fArrayOnly = " & CStr(fArrayOnly), glngUserID, Now, "Start Update of GradeList")
    
   OpenGradeDatabase     ' ltas 8.29.2006
    If dbGrade = 0 Then
        MsgBox "Unable to open Grade database"
        Exit Sub
    End If
    ErrorCheck (dbGrade)
    PDBSetSort dbGrade, "DisplayOrder,Grade"
    ErrorCheck (dbGrade)
    
    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If
    
    If fFirstGradeLoad = True Then
        I = -1
    Else
        ReDim garyGrade(25, 0)
        ReDim gudtGrade(0)
        
        I = 0
        fFirstGradeLoad = True
    End If
    
    
    PDBMoveFirst dbGrade
    If PDBEOF(dbGrade) = True Then
        '     '     Call WriteEventLog(0, "UPDATEGRADELIST", "0 Count Check", glngUserID, Now, "There were no records in teh database prior to going through the loop????")
    End If
    
    While Not PDBEOF(dbGrade)
        ReadGradeRecord udt
        If fArrayOnly = False Then
            cbo.AddItem strGrade
            cbo.ItemData(cbo.NewIndex) = udt.GradeID
        End If
        
        If I <> -1 Then
            If I = 0 And garyGrade(0, 0) = "" Then
                'Do Nothing..noredim needed
            Else
                I = I + 1
                ReDim Preserve garyGrade(25, I)
                ReDim Preserve gudtGrade(I)
            End If
            'Write event log for trouble shooting
            '     '     Call WriteEventLog(0, "UPDATEGRADELIST", "AddtoArray-CodeBlock", glngUserID, Now, "Adding Records to Array")
            
            garyGrade(0, I) = udt.GradeID
            garyGrade(1, I) = udt.Grade
            
            garyGrade(2, I) = udt.HHAID
            garyGrade(3, I) = udt.GradingUse
            garyGrade(4, I) = udt.DisplayOrder
            garyGrade(5, I) = udt.Cant
            garyGrade(6, I) = udt.Width
            garyGrade(7, I) = udt.thickness
            garyGrade(8, I) = udt.CantMask
            garyGrade(9, I) = udt.ReceivingUse
            'Downloading the GradeGroup as the Grouping Field Now so don't have to update pdb structure
            garyGrade(10, I) = udt.Grouping
            'Added 8;/12/16 when pdb structure changed:
            garyGrade(11, I) = udt.GradeType
            garyGrade(12, I) = udt.HuskyGradeID
            garyGrade(13, I) = udt.SilvaTechGradeID
            garyGrade(14, I) = udt.HHSortOrder
            garyGrade(15, I) = udt.SOAID
            garyGrade(16, I) = udt.Protected
            garyGrade(17, I) = udt.HHActive
            garyGrade(18, I) = udt.GradeGroupFlag
            garyGrade(19, I) = udt.GGID
            garyGrade(20, I) = udt.HHAIDNum
            garyGrade(21, I) = udt.GradeX1
            garyGrade(22, I) = udt.GradeX2
            garyGrade(23, I) = udt.GradeX3
            garyGrade(24, I) = udt.GradeX4
            garyGrade(25, I) = udt.GradeX5
            
            
            'New Grade array of UDT - Global
            gudtGrade(I).GradeID = udt.GradeID
            gudtGrade(I).Grade = udt.Grade
            gudtGrade(I).Grouping = udt.Grouping
            gudtGrade(I).HHAID = udt.HHAID
            gudtGrade(I).GradingUse = udt.GradingUse
            gudtGrade(I).DisplayOrder = udt.DisplayOrder
            gudtGrade(I).Cant = udt.Cant
            gudtGrade(I).Width = udt.Width
            gudtGrade(I).thickness = udt.thickness
            gudtGrade(I).CantMask = udt.CantMask
            gudtGrade(I).ReceivingUse = udt.ReceivingUse
            gudtGrade(I).GradeType = udt.GradeType
            gudtGrade(I).HuskyGradeID = udt.HuskyGradeID
            gudtGrade(I).SilvaTechGradeID = udt.SilvaTechGradeID
            gudtGrade(I).HHSortOrder = udt.HHSortOrder
            gudtGrade(I).SOAID = udt.SOAID
            gudtGrade(I).Protected = udt.Protected
            gudtGrade(I).HHActive = udt.HHActive
            gudtGrade(I).GradeGroupFlag = udt.GradeGroupFlag
            gudtGrade(I).GGID = udt.GGID
            gudtGrade(I).HHAIDNum = udt.HHAIDNum
            gudtGrade(I).GradeX1 = udt.GradeX1
            gudtGrade(I).GradeX2 = udt.GradeX2
            gudtGrade(I).GradeX3 = udt.GradeX3
            gudtGrade(I).GradeX4 = udt.GradeX4
            gudtGrade(I).GradeX5 = udt.GradeX5


            
        End If
        
        PDBMoveNext dbGrade
    Wend
    PDBClose dbGrade
        
    If fArrayOnly = False Then cbo.ListIndex = 0
    Exit Sub
ErrorHandler:
    MsgBox "Error in UpdateGradeList " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub

Public Function GetGradeData(strSearchString As String, strSearchType As String, Optional strReturnSingle As String, _
                            Optional lngReturnID As Long, _
                            Optional strReturnAID As String, _
                            Optional strReturnName As String) As String
    
    Dim I As Integer
    Dim fFound As Boolean
    
On Error GoTo ErrorHandler

    GetGradeData = "-1,INVALID, ,-1,0,0,0,0,"
    For I = 0 To UBound(garyGrade, 2)
        
        If SC(strSearchType, "AID") Or SC(strSearchType, "HHAID") Then
            If Trim(UCase(garyGrade(2, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "Grade") = True Or SC(strSearchType, "GRADENAME") = True Or SC(strSearchType, "GRADEDESC") = True Or SC(strSearchType, "NAME") = True Or SC(strSearchType, "DESC") = True Then
            If Trim(UCase(garyGrade(1, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "GradingUse") Then
            If Trim(UCase(garyGrade(3, I))) = (Trim(UCase(strSearchString))) Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "GROUP") = True Or SC(strSearchType, "GROUPING") = True Then
            If Trim(UCase(garyGrade(10, I))) = (Trim(UCase(strSearchString))) Then
                'Match Found
                fFound = True
            End If
        Else
            If Trim(UCase(garyGrade(0, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        End If
        
        If fFound = True Then
            If SC(strReturnSingle, "") = True Then
                GetGradeData = garyGrade(0, I) & "," & garyGrade(1, I) & "," & garyGrade(2, I) & "," & garyGrade(3, I) & "," & garyGrade(4, I) & "," & garyGrade(5, I) & "," & garyGrade(6, I) & "," & garyGrade(7, I) & "," & garyGrade(8, I) & "," & garyGrade(9, I) & "," & garyGrade(10, I)
                
            ElseIf SC(strReturnSingle, "ID") Then
                GetGradeData = garyGrade(0, I)
                
            ElseIf SC(strReturnSingle, "CANT") Then
                GetGradeData = garyGrade(5, I)
                
            ElseIf SC(strReturnSingle, "GRADE") = True Then
                GetGradeData = garyGrade(1, I)
                
            ElseIf SC(strReturnSingle, "HHAID") = True Then
                GetGradeData = garyGrade(2, I)
                
            ElseIf SC(strReturnSingle, "SORTORDER") = True Then
                GetGradeData = garyGrade(4, I)
                
            ElseIf SC(strReturnSingle, "CANTMASK") Then
                GetGradeData = garyGrade(8, I)
                
            ElseIf SC(strReturnSingle, "GROUPING") = True Or SC(strReturnSingle, "GROUP") Then
                GetGradeData = garyGrade(10, I)
            
            'New Additions use the gudtGrade() array as it's a userdefined type/easier to use and loaded simultaneously with this array
            ElseIf SC(strReturnSingle, "HuskyGradeID") = True Or SC(strReturnSingle, "HUSKY") = True Then
                GetGradeData = gudtGrade(I).HuskyGradeID
            
            ElseIf SC(strReturnSingle, "SilvaTechGradeID") = True Or SC(strReturnSingle, "SILVAID") = True Or SC(strReturnSingle, "SILVA") = True Then
                GetGradeData = gudtGrade(I).SilvaTechGradeID
            
            ElseIf SC(strReturnSingle, "GGID") = True Or SC(strReturnSingle, "GRADEGROUPID") = True Then
                GetGradeData = CStr(gudtGrade(I).GGID)
            
            ElseIf SC(strReturnSingle, "GGFlag") = True Or SC(strReturnSingle, "GradeGroupFlag") = True Then
                GetGradeData = CStr(gudtGrade(I).GradeGroupFlag)
            
            ElseIf SC(strReturnSingle, "GradeX1") = True Or SC(strReturnSingle, "X1") = True Then
                GetGradeData = CStr(gudtGrade(I).GradeX1)
            
            ElseIf SC(strReturnSingle, "GradeX2") = True Or SC(strReturnSingle, "X2") = True Then
                GetGradeData = CStr(gudtGrade(I).GradeX2)
            
            ElseIf SC(strReturnSingle, "GradeX3") = True Or SC(strReturnSingle, "X3") = True Then
                GetGradeData = CStr(gudtGrade(I).GradeX3)
            
            ElseIf SC(strReturnSingle, "GradeX4") = True Or SC(strReturnSingle, "X4") = True Then
                GetGradeData = CStr(gudtGrade(I).GradeX4)
            
            ElseIf SC(strReturnSingle, "GradeX5") = True Or SC(strReturnSingle, "X5") = True Then
                GetGradeData = CStr(gudtGrade(I).GradeX5)
            Else
                MsgBox "Return Type for Grade Lookup Type=" & strReturnSingle & " - Is Not Found - Returning Full grade string! Contact eLIMBS at 740.401.0720 "
            End If
            
            Exit For
        End If
        
    Next
        
    If fFound = True Then
        strReturnAID = gudtGrade(I).HHAID
        strReturnName = gudtGrade(I).Grade
        lngReturnID = gudtGrade(I).GradeID
    Else
        strReturnAID = "NA"
        strReturnName = "NOT FOUND"
        lngReturnID = -1
        
        If fFound = False And strReturnSingle = "ID" Then
            GetGradeData = "-1"
            
        ElseIf fFound = False And SC(strReturnSingle, "SortOrder") = True Then
            GetGradeData = "999"
            
        ElseIf fFound = False And SCInList(strReturnSingle, "GROUP,GROUPING,GGID,GRADEGROUPID,GRADEGROUPFLAG,GGFLAG") = True Then
            GetGradeData = "-1"
            
        ElseIf fFound = False And (SC(strReturnSingle, "SilvaTechGradeID") = True Or SC(strReturnSingle, "SILVAID") = True Or SC(strReturnSingle, "SILVA") = True) Then
            GetGradeData = "INVALID"
            
        End If
    End If
    
    Exit Function

ErrorHandler:
    MsgBox "Error in GetGradeData(SearchString=" & strSearchString & " Search Type=" & strSearchType & " Return Single=" & strReturnSingle & " " & vbCrLf & Err.Number & "-" & Err.Description
    Exit Function
    
End Function
Public Sub UpdateLengthList(cbo As AFComboBox, Optional strStatus As String, Optional fArrayOnly As Boolean)

    Dim strText As String
    Dim strText2 As String
    Dim lngID As Long
    Dim fAdd As Boolean
    Dim I As Integer
    Dim strHHAID As String
    Dim intMin As Integer
    Dim intMax As Integer
    Dim intDisplayOrder As Integer


    Dim udt As tLengthRecord

On Error GoTo ErrorHandler

    OpenLengthDatabase     ' New Module for New Structure Created 20150619 by PCC
        '       keeping all the old arry "stuff" for garylength fields (only ones that existed in 9/2004 when it was created for backwards code compatibility
    
    If dbLength = 0 Then
        MsgBox "Unable to open Length database"
        Exit Sub
    End If

    PDBSetSort dbLength, "ProductionStatusType,LengthDisplayOrder,LengthName"

    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If

    ReDim gudtLenNew(0)
    gudtLenNew(0).LengthID = -1


    ReDim garyLength(tLengthDatabaseFields.LenMaxCM, 0)
    I = 0

    PDBMoveFirst dbLength

    'Check to see if Only Air Dried or Kiln Dried Status' should be used
    If strStatus <> "" And (Left(UCase(strStatus), 2) = "AI" Or Left(UCase(strStatus), 2) = "AD") Then
        strStatus = "ROUGH GRADED"
    ElseIf strStatus <> "" Then
        strStatus = "FINAL GRADED"
    End If

    While Not PDBEOF(dbLength)
        If UBound(gudtLenNew) = 0 And gudtLenNew(0).LengthID <= 0 Then
            ' do nothing, first record slot was dimmed above
        Else
            ReDim Preserve gudtLenNew(UBound(gudtLenNew) + 1)
        End If

        PDBReadRecord dbLength, VarPtr(gudtLenNew(UBound(gudtLenNew)))

        fAdd = False
        If strStatus = "" Then
            fAdd = True
        Else
            If SC(gudtLenNew(UBound(gudtLenNew)).ProductionStatusType, strStatus) = True Then
                fAdd = True
            End If
        End If

        If fAdd = True Then

            If fArrayOnly = False Then
                cbo.AddItem gudtLenNew(UBound(gudtLenNew)).LengthName & "-" & gudtLenNew(UBound(gudtLenNew)).ProductionStatusType
                cbo.ItemData(cbo.NewIndex) = gudtLenNew(UBound(gudtLenNew)).LengthID
            End If

            If I = 0 And garyLength(0, 0) = "" Then
                'Do Nothing..noredim needed
            Else
                I = I + 1
                ReDim Preserve garyLength(tLengthDatabaseFields.LenMaxCM, I)
            End If

            'Only doing this for backward compatibility, all new calls will use this new function anyway and it won't matter
            'but just incase it's manually split/looked up elsewhere, populating the array with the fields
            'from 2004 only, not new fields 2015. PCC 20150619
            
            garyLength(0, I) = gudtLenNew(UBound(gudtLenNew)).LengthID
            garyLength(1, I) = gudtLenNew(UBound(gudtLenNew)).LengthName
            garyLength(2, I) = gudtLenNew(UBound(gudtLenNew)).ProductionStatusType
            garyLength(3, I) = gudtLenNew(UBound(gudtLenNew)).HHAID
            garyLength(4, I) = CStr(gudtLenNew(UBound(gudtLenNew)).MinLength)
            garyLength(5, I) = CStr(gudtLenNew(UBound(gudtLenNew)).MaxLength)
            garyLength(6, I) = CStr(gudtLenNew(UBound(gudtLenNew)).LengthDisplayOrder)
            
            If SC(tcu(gudtLenNew(I).LengthUnit), "FEET") = True Then
                gudtLenNew(I).LengthUnit = "FT"
            ElseIf SC(tcu(gudtLenNew(I).LengthUnit), "INCHES") = True Then
                gudtLenNew(I).LengthUnit = "IN"
            End If
            
            
            If CDbl(gudtLenNew(I).CalcFactor) <= 0 Then
                If IsNumeric(Replace(Replace(gudtLenNew(I).LengthName, "'", ""), """", "")) = True Then
                    gudtLenNew(I).CalcFactor = CDbl(Replace(Replace(gudtLenNew(I).LengthName, "'", ""), """", ""))
                End If
            End If
        End If

        PDBMoveNext dbLength
    Wend

    PDBClose dbLength
    dbLength = 0

    If fArrayOnly = False Then cbo.ListIndex = 0
    Exit Sub
ErrorHandler:
    dbLength = 0
    MsgBox "Error in UpdateLengthList " & Err.Number & "-" & Err.Description
    Exit Sub
End Sub
Public Function GetLengthData(strSearchString As String, strSearchType As String, Optional strReturnField As String, _
    Optional strLenMin As String, Optional strLenMax As String, Optional ByRef intReturnIndex_gudtLenNew As Integer) As String
    Dim I As Integer
    Dim fFound As Boolean
On Error GoTo ErrorHandler
    
    'Keeping the OLD "TextStuff/ARRAYStuff" For the Old Code that is already in the system using it
    'updating /adding all the new fields from orignial design of 9/2004 to current of 6/19/2015 by PCC
    GetLengthData = "-1,INVALID, , ,0,0,0"
    
    For I = 0 To UBound(gudtLenNew)
        If SC(strSearchType, "LengthName") = True Or SC(strSearchType, "NAME") = True Then
            If Trim(UCase(gudtLenNew(I).LengthName)) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "ProductionStatusType") Then
            If Trim(UCase(gudtLenNew(I).ProductionStatusType)) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "HHAID") Then
            If Trim(UCase(gudtLenNew(I).HHAID)) = Trim(UCase(strSearchString)) Then
                'Match Found
                strLenMin = CStr(gudtLenNew(I).MinLength)
                strLenMax = CStr(gudtLenNew(I).MaxLength)
                fFound = True
            End If
        Else 'ID
            If Trim(UCase(gudtLenNew(I).LengthID)) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        End If
        
        If fFound = True Then
            If strReturnField = "" Then
                GetLengthData = gudtLenNew(I).LengthID & "," & gudtLenNew(I).LengthName & "," & gudtLenNew(I).ProductionStatusType & "," & _
                        "" & gudtLenNew(I).HHAID & "," & gudtLenNew(I).MinLength & "," & gudtLenNew(I).MaxLength & "," & gudtLenNew(I).LengthDisplayOrder
            ElseIf Trim(UCase(strReturnField)) = "LENGTHNAME" Then
                GetLengthData = gudtLenNew(I).LengthName
            ElseIf Trim(UCase(strReturnField)) = "ID" Then
                GetLengthData = gudtLenNew(I).LengthID
            ElseIf Trim(UCase(strReturnField)) = "HHAID" Then
                GetLengthData = gudtLenNew(I).HHAID
            ElseIf SC(strReturnField, "CalcFactor") = True Then
                GetLengthData = gudtLenNew(I).CalcFactor
            End If
            
            intReturnIndex_gudtLenNew = I
            Exit For
        ElseIf fFound = False Then
            'This is currently being used for printing tags so don't send back error messages
                'because blank is an ok entry
            If strReturnField = "" Then
                'just continue...don't update the value from the intial error settings
            ElseIf Trim(UCase(strReturnField)) = "LENGTHNAME" Then
                GetLengthData = ""
            ElseIf SC(strReturnField, "CalcFactor") = True Then
                GetLengthData = "-1"
            End If
        End If
    
    Next
    
    If fFound = False Then intReturnIndex_gudtLenNew = -1 ' always return this for the optional index of the array being returned for the
        'new type array added in 6/2015 by PCC
        
    If fFound = False And (SC(strReturnField, "CalcFactor") = True Or strReturnField = "ID") Then
        GetLengthData = "-1"
    End If
    
    Exit Function
       
ErrorHandler:
    MsgBox "Error in GetLengthData " & Err.Number & "-" & Err.Description
    Exit Function
       
End Function


Public Sub GetLengthRangeData(intMax As Integer, intMin As Integer, ByRef lngLengthID As Long, _
                            ByRef strLength As String, strType As String, _
                            Optional strLengthAID As String, Optional fLenghtInLenches As Boolean, Optional strLengthUnit As String, _
                            Optional dblMax As Double, Optional dblMin As Double, Optional fUseDoubles As Boolean)
    Dim aryIndex() As Integer
    Dim J As Integer
    Dim I As Integer
    Dim dblDiff As Double
    Dim dblBestDiff As Double
    Dim intFound_BestIndex As Integer
    Dim fSkip As Boolean
    Dim lngLengthIDTemp As Integer
    Dim strLengthTemp As String
    
    
    
    '*********************************IMPORTANT READ IF ANY ISSUES *********************************
    '*************                                                 *********************************
    '*************                                                 *********************************
    '*************                                                 *********************************
    '*************              major change by PCC                *********************************
    '*************               1/1/2017                          *********************************
    '*************               w/o Thorough testing/review       *********************************
    '************* If issues rever to file from source safe prior to 1/17/2017 *** TALK TO PCC *****
    '************* BUT DO IT IMMEDIATELY IF PRODUCTION/BUNLE ISSUE AT ANY CLIEN OTHER THAN STELLA JONES
    
    
On Error GoTo ErrorHandler
    '***************NOT UPDATED TO USE NEW gudtLenNew() of tblLengthDatabaseRecordType added with new structure created 6/19/2015
        'by PCC - This module would still benefit from being updated
    
    lngLengthIDTemp = lngLengthID
    strLengthTemp = strLength
    lngLengthID = 0
    strLength = ""
    
    If fUseDoubles = True Then
        intMin = CInt(dblMin) 'default to integer conversion of dblmin
        intMax = CInt(dblMax) 'default to integer conversion of dblMax
    
        If InStr(CStr(dblMin), ".") > 0 Or InStr(CStr(dblMax), ".") > 0 Then
            intMin = GetAvgLengthMinMax(CStr(dblMin), "MIN")
            intMax = GetAvgLengthMinMax(CStr(dblMax), "MAX")
        End If
    Else
        'set double values = to integer values
        dblMin = CDbl(intMin)
        dblMax = CDbl(intMax)
        
    End If
    
    
        
    'If they are decimal values then change from the above to the below using speecifically defined function getavglengthminmax ...
    
    
On Error GoTo ErrorHandler

    ReDim aryIndex(0)
    aryIndex(0) = -1 ' this index will eliminate options for length ranges
    For I = LBound(garyLength, 2) To UBound(garyLength, 2)
        Dim fValidLength As Boolean
        
        fValidLength = True
        
        If fLenghtInLenches = True Then
            If SC(gudtLenNew(I).LengthUnit, "IN") = True Then
                fValidLength = True
            Else
                fValidLength = False
            End If
        Else
            fValidLength = True
        End If
        
        If fValidLength = True Then
            If strType = "" Then
                'do nothing..compare against all
            ElseIf strType <> UCase(garyLength(2, I)) Then
                'exclude this on based on the type
                ReDim Preserve aryIndex(UBound(aryIndex) + 1)
                aryIndex(UBound(aryIndex)) = I
                
            ElseIf CDbl(garyLength(4, I)) > dblMin Or CDbl(garyLength(5, I)) < dblMax Then 'if the minimum in the range is greater than
                'the actual min lenght..it's out or the maximum of the range is less than
                'the actual max length
                ReDim Preserve aryIndex(UBound(aryIndex) + 1)
                aryIndex(UBound(aryIndex)) = I
            Else
                'Do Nothing
                Debug.Print I
            End If
            
            Debug.Print "Length Name = " & garyLength(1, I) & "   HHAID=" & garyLength(3, I) & "    Min-Max = " & garyLength(4, I) & "-" & CDbl(garyLength(5, I))
            If IsNumeric(garyLength(4, I)) = True And IsNumeric(garyLength(5, I)) = True Then
                
                If CDbl(garyLength(4, I)) = dblMin And CDbl(garyLength(5, I)) = dblMax Then 'this is it
                    lngLengthID = CInt(garyLength(0, I))
                    strLength = CStr(garyLength(1, I))
                    strLengthAID = CStr(garyLength(3, I))
                    Exit For
                End If
            End If
        End If
    Next
    'Check to see if an exact match was found
    'If it is an exact match, select it
        
    If lngLengthID <> 0 And strLength <> "" Then
        Exit Sub
    End If
    
    '*****************************************Changing this back to the way it was essentially by using the
    '*****************************************new variables I added when I made the change to allow decimal/exact match values
    '*****************************************replacing intMIn/intMax below with intMin/intMax
    
    
    '*********************THE BELOW DOESNT USE THE DECIMAL VALUES LIKE THE EXACT MATCHES ABOEV....IT"S THE WAY IT ALWAYS WAS IF I DID CORRECTLY
    '************** NOTE *** ** DONE VERY QUICKLY >>>IF NEEDED REVERT BACK TO PRe 1/1/2017 version from source safe and update client IMMEDIATELY
    '*************** IF PCC CANNOT BE REACHED.....THEN KEEP TRYING TO REACH PCC EVERY 3 Hour Max between tries until contact achieved...don't
    '*************** delay fixing client .....because can't get PCC...DO AS ABOVE SAYS.
    
    'If no exact match was found, try and find the best fit
    dblBestDiff = 100
    intFound_BestIndex = -1
    For I = LBound(garyLength, 2) To UBound(garyLength, 2)
        fSkip = False
        For J = 0 To UBound(aryIndex)
            If aryIndex(J) = I Then
                fSkip = True
                Exit For
            End If
        Next
        
        If fSkip = True Then 'this on is already excluded ... go to the next I
            'Do Nothing
        Else
            dblDiff = Abs(intMin - CInt(garyLength(4, I))) + Abs(CInt(garyLength(5, I)) - intMax)
            
            If intMin >= CInt(garyLength(4, I)) And intMax <= CInt(garyLength(5, I)) And dblDiff < dblBestDiff Then
                dblBestDiff = dblDiff
                intFound_BestIndex = I
                
            ElseIf dblDiff = dblBestDiff Then 'compare for an exact min or max match
                If intMin - CInt(garyLength(4, I)) = 0 Then
                    dblBestDiff = dblDiff
                    intFound_BestIndex = I
                    
                ElseIf CInt(garyLength(5, I)) - intMax = 0 Then
                    dblBestDiff = dblDiff
                    intFound_BestIndex = I
                    
                End If
            End If
        End If
    Next
    
    If intFound_BestIndex = -1 Then
        'no matchs
        lngLengthID = lngLengthIDTemp
        strLength = strLengthTemp
    Else
        lngLengthID = CLng(garyLength(0, intFound_BestIndex))
        strLength = garyLength(1, intFound_BestIndex)
        strLengthAID = garyLength(3, intFound_BestIndex)
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error In GetLengthRangeData No Matching Length Found"
    Exit Sub
    
End Sub
Public Sub UpdateproductionstatusList(cbo As AFComboBox, Optional ByVal strStatus As String, Optional fArrayOnly As Boolean)
'LTAS 4-6-2008
    Dim strType As String
    Dim strText As String
    Dim strTextCheck As String
    Dim I As Integer
    Dim strHHAID As String
    Dim udt As tProductionStatusRecord
    
  '****LONG STANDING ISSUE IDENTIFIED:  Dim lngID As Long. THE DB FIELD IS AN INTEGER! Found 4-7-2008
  
    Dim lngID As Long
    Dim fAdd As Boolean
    
   OpenProductionStatusDatabase     ' ltas 8.29.2006
    If dbProductionStatus = 0 Then
        MsgBox "Unable to open productionstatus database"
        Exit Sub
    End If
    
    PDBSetSortFields dbProductionStatus, 1
    
    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If
    
    PDBMoveFirst dbProductionStatus
    


    If strStatus = "CANTS/OUTS" Then
        strStatus = "CANTS/OUTS"
        'repetitive...but just making the point
    Else
        strStatus = ""
    End If
    
    I = -1
    While Not PDBEOF(dbProductionStatus)
        fAdd = False
        
        If strStatus = "" Then
            fAdd = True
        Else
            PDBGetField dbProductionStatus, 2, strTextCheck
            If UCase(strTextCheck) = strStatus Then
                fAdd = True
            End If
        End If
        
        If fAdd = True Then
            I = I + 1
            If I = 0 Then
                ReDim garyStatus(4, 0)
            Else
                ReDim Preserve garyStatus(4, I)
            End If
            
            PDBReadRecord dbProductionStatus, VarPtr(udt)
            
            If fArrayOnly = False Then
                cbo.AddItem strText
                cbo.ItemData(cbo.NewIndex) = udt.StatusID
            End If
            
            garyStatus(0, I) = udt.StatusID
            garyStatus(1, I) = udt.ProductionStatusName
            garyStatus(2, I) = udt.ProductionStatusName
            garyStatus(3, I) = udt.GreenDry
            garyStatus(4, I) = udt.PSAID
            
        End If
        PDBMoveNext dbProductionStatus
    Wend
    CloseProductionStatusDatabase
        
    If fArrayOnly = False Then cbo.ListIndex = 0

End Sub
Public Function GetStatusData(strSearchString As String, strSearchType As String, Optional strReturnField As String) As String
    Dim I As Integer
    Dim fFound As Boolean
    
    
    GetStatusData = "-1,INVALID,,,NA "
    
    For I = 0 To UBound(garyStatus, 2)
        If strSearchType = "STATUS" Then
            If Trim(UCase(garyStatus(1, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf strSearchType = "PSAID" Then
            If Trim(UCase(garyStatus(4, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
            
        Else
            If Trim(UCase(garyStatus(0, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        End If
        
        If fFound = True Then
            
            If strReturnField = "STATUS" Or strReturnField = "NAME" Then
                GetStatusData = garyStatus(1, I)
                Exit For
            ElseIf strReturnField = "ID" Then
                GetStatusData = garyStatus(0, I)
                Exit For
            ElseIf SC(strReturnField, "PSAID") = True Or SC(strReturnField, "HHAID") = True Then
                GetStatusData = garyStatus(4, I)
                Exit For
            Else
                GetStatusData = garyStatus(0, I) & "," & garyStatus(1, I) & "," & garyStatus(2, I) & "," & garyStatus(3, I) & "," & garyStatus(4, I)
                Exit For
            End If
        End If
                    
    Next
    
    If Trim(strReturnField) = "" And fFound = False Then
        'nothing
    ElseIf Trim(strReturnField) <> "" And fFound = False Then
        GetStatusData = "-1"
    End If
        
    Exit Function
       
End Function

Public Sub UpdateOrgList(cbo As AFComboBox, strOrgType As String, fArrayOnly As Boolean)
    
    Dim strText As String
    Dim strText2 As String
    Dim lngID As Long
    Dim fAdd As Boolean
    Dim strType As String
    Dim strHHAID As String
    Dim I As Integer
    Dim strOrgPhone As String
    
    
    
  OpenOrgDatabase     ' ltas 8.29.2006
    If dbOrg = 0 Then
        MsgBox "Unable to open Length database"
        Exit Sub
    End If
    
    PDBSetSort dbOrg, "OrgType,OrgName"
    
    If fArrayOnly = False Then
        cbo.AddItem ""
        cbo.ItemData(0) = 0
    End If
    
    I = -1
    
    PDBMoveFirst dbOrg
    If PDBEOF(dbOrg) = True Then
        ReDim garyOrg(4, 0)
        garyOrg(0, 0) = 0
    End If
    
    While Not PDBEOF(dbOrg)
        fAdd = False
        If strOrgType <> "" Then
            PDBGetField dbOrg, 4, strText2
            If UCase(strText2) <> UCase(strOrgType) Then
                'Add the field
                fAdd = True
            End If
        Else
            fAdd = True
        End If
        
        PDBGetField dbOrg, 0, lngID
        PDBGetField dbOrg, tOrgDatabaseFields.OrgName_Field, strText
        PDBGetField dbOrg, 7, strType
        PDBGetField dbOrg, 2, strHHAID
        PDBGetField dbOrg, 3, strOrgPhone
        
        
        If fArrayOnly = False And fAdd = True Then
            cbo.AddItem strText
            cbo.ItemData(cbo.NewIndex) = lngID
        End If
        
        If fArrayOnly = True Then
            If I = -1 Then
                I = I + 1
                ReDim garyOrg(4, 0)
            Else
                I = I + 1
                ReDim Preserve garyOrg(4, I)
            End If
            
            garyOrg(0, I) = lngID
            garyOrg(1, I) = strText
            garyOrg(2, I) = strType
            garyOrg(3, I) = strHHAID
            garyOrg(4, I) = strOrgPhone
        End If
        
        PDBMoveNext dbOrg
        
    Wend
    
    PDBClose dbOrg
    If fArrayOnly = False Then cbo.ListIndex = 0

End Sub

Public Function GetOrgData(strSearchString As String, strSearchType As String, Optional strOrgType As String, _
                            Optional strReturnField As String, Optional lngReturnID As Long, _
                            Optional strReturnAID As String, Optional strReturnName As String) As String
    Dim I As Integer
    Dim fFound As Boolean
    Dim strSearchTest As String
 On Error GoTo ErrorHandler
 
    GetOrgData = "-1,INVALID, , "
    If SC(strReturnField, "NAME") = True Then
        GetOrgData = "NA"
    End If
    
    For I = 0 To UBound(garyOrg, 2)
        'This defaults the search test to whatever is currently in the array position if no specific
            'instructions were sent into the OrgType Option
        If strOrgType = "" Then
            strSearchTest = UCase(garyOrg(2, I))
        Else
            strSearchTest = UCase(strOrgType)
        End If
        
        If SC(strSearchType, "VENDORNAME") Or SC(strSearchType, "VENDORDESC") Or SC(strSearchType, "CARRIERAID") Or SC(strSearchType, "CARRIERNAME") _
                Or SC(strSearchType, "OrgName") Or SC(strSearchType, "NAME") Or SC(strSearchType, "DESC") Then
            
            If Trim(UCase(garyOrg(1, I))) = Trim(UCase(strSearchString)) And UCase(garyOrg(2, I)) = strSearchTest Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "HHAID") Or SC(strSearchType, "AID") Then
            If Trim(UCase(garyOrg(3, I))) = Trim(UCase(strSearchString)) Then 'And UCase(garyOrg(2, I)) = strSearchTest Then
                'Match Found
                fFound = True
            End If
        ElseIf SC(strSearchType, "ID") Or SC(strSearchType, "ORGID") Then
            If Trim(UCase(garyOrg(0, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        ElseIf (UCase(garyOrg(2, I)) = strSearchTest) Or UCase(garyOrg(2, I)) = "" Then 'ID
            If Trim(UCase(garyOrg(0, I))) = Trim(UCase(strSearchString)) Then
                'Match Found
                fFound = True
            End If
        End If
        
        If fFound = True Then
            If SCInList(strReturnField, "NAME,ORGNAME,VENDORNAME,DESC,CARRIERNAME") = True Then
                GetOrgData = garyOrg(1, I)
            ElseIf SCInList(strReturnField, "HHAID,AID,ABBREV") = True Then
                GetOrgData = garyOrg(3, I)
            ElseIf SC(strReturnField, "ID") = True Then
                GetOrgData = garyOrg(0, I)
            Else
                GetOrgData = garyOrg(0, I) & "," & garyOrg(1, I) & "," & garyOrg(2, I) & "," & garyOrg(3, I) & "," & garyOrg(4, I)
            End If
            lngReturnID = CLng(garyOrg(0, I))
            strReturnAID = garyOrg(3, I)
            strReturnName = garyOrg(1, I)
            
            Exit For
        End If
                    
    Next
    
    If fFound = False Then
        If SCInList(strReturnField, "NAME,ORGNAME,VENDORNAME,DESC,CARRIERNAME") = True Then
            GetOrgData = "NOTFOUND"
        ElseIf fFound = False And SCInList(strReturnField, "HHAID,AID,ABBREV") = True Then
            GetOrgData = "NA"
        ElseIf fFound = False And SC(strReturnField, "ID") = True Then
            GetOrgData = "-1"
        Else
            'return the default/old methodology/string from above
            GetOrgData = "-1,INVALID, , "
        End If
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error in GetOrgData : String=" & strSearchString & " SearchType=" & strSearchType & vbCrLf & Err.Number & "-" & Err.Description
    If SC(strReturnField, "ID") = True Then
        GetOrgData = "-1"
    Else
        GetOrgData = "ERROR-" & Err.Number
    End If
    Exit Function
        
    
End Function
Public Function GetLocData(lngLocID As Long, strLocAID As String, strReturnField As String, _
                            Optional lngReturnID As Long, Optional _
                            strReturnAID As String, Optional strReturnName As String) As String
    Dim I As Integer
    Dim fFound As Boolean
    
On Error GoTo ErrorHandler
    
    
    If SC(strReturnField, "ID") = True Then
        GetLocData = -1
    Else
        GetLocData = "ERROR"
    End If
    
    For I = 0 To UBound(garyLoc)
        If lngLocID <= 0 And Trim(UCase(strLocAID)) = Trim(UCase(garyLoc(I).LocAID)) Then
            If Trim(UCase(strReturnField)) = "LOCDESC" Then
                GetLocData = garyLoc(I).LOCDesc
                fFound = True
                Exit For
            ElseIf SC(strReturnField, "ID") = True Or SC(strReturnField, "LOCID") = True Then
                GetLocData = garyLoc(I).LOCID
                lngLocID = garyLoc(I).LOCID
                fFound = True
                Exit For
            ElseIf SCInList(strReturnField, "NAME,LOCNAME,DESC,DESCRIPTION,LOCDESC") = True Then
                GetLocData = garyLoc(I).LOCDesc
                fFound = True
                Exit For
            ElseIf SC(strReturnField, "LocX1") = True Then 'User Related to Location for Filtering on Loc F1 / Handheld
                GetLocData = garyLoc(I).LocX1
                fFound = True
                Exit For
            ElseIf SC(strReturnField, "LocX2") = True Then 'As of 2/6/2017 LocJoinAID1
                GetLocData = garyLoc(I).LocX2
                fFound = True
                Exit For
            ElseIf SC(strReturnField, "LocX3") = True Then 'As of 2/6/2017 LocJoinAID2
                GetLocData = garyLoc(I).LocX3
                fFound = True
                Exit For
            ElseIf SC(strReturnField, "LocX4") = True Then 'As of 2/6/2017 LocJoinAID3
                GetLocData = garyLoc(I).LocX4
                fFound = True
                Exit For
            End If
        End If
    Next
    
    If fFound = False Then
        lngReturnID = -1
        strReturnAID = strLocAID
        strReturnName = "NOT FOUND"
    Else
        lngReturnID = garyLoc(I).LOCID
        strReturnAID = garyLoc(I).LocAID
        strReturnName = garyLoc(I).LOCDesc
    End If
    
    If SC(strReturnField, "ID") = True And fFound = False Then
        GetLocData = "-1"
    ElseIf fFound = False Then
        GetLocData = strLocAID
    ElseIf fFound = True Then
        'getlocdata value was set in the loop above based upon the return field/correct return value
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error in GetLocData lngLocID=" & lngLocID & " strLocAID=" & strLocAID & " ReturnField=" & strReturnField & vbCrLf & Err.Number & "-" & Err.Description
    If SC(strReturnField, "ID") = True Then
        GetLocData = "-1"
    Else
        GetLocData = "ERROR"
    End If
    Exit Function
End Function
Public Function GetMacData(strMacAID As String, strReturnField As String) As String
    Dim I As Integer
    Dim fFound As Boolean
    
    fFound = False
    GetMacData = "ERROR"
    
    For I = 0 To UBound(garyMAC)
        If tcu(strMacAID) = tcu(garyMAC(I).MACAID) Then
            If Trim(UCase(strReturnField)) = "ID" Then
                fFound = True
                GetMacData = garyMAC(I).MACID
                Exit For
            ElseIf tcu(strReturnField) = "HHAID" Then
                fFound = True
                GetMacData = garyMAC(I).MACAID
                Exit For
            Else
                GetMacData = garyMAC(I).MACAID
                fFound = True
                Exit For
            End If
        End If
    Next
    
    If fFound = False And tcu(strReturnField) = "ID" Then
        GetMacData = "-1"
    ElseIf fFound = False Then
        GetMacData = "ERROR"
    End If
    
    Exit Function
End Function

Public Function GetOrderData(strOrderAID As String, lngReturnField As Long) As String
    Dim I As Integer
    Dim fFound As Boolean
    Dim udt As tOHRecord
    
    OpenOHDatabase
    PDBSetSortFields dbOH, tOHDatabaseFields.OHAID
    PDBMoveFirst dbOH
    
    PDBFindRecordByField dbOH, tOHDatabaseFields.OHAID, strOrderAID
    
    If PDBGetLastError(dbOH) <> 0 Then
        GetOrderData = "-1"
        CloseOHDatabase
        dbOH = 0
    Else
        PDBReadRecord dbOH, VarPtr(udt)
        If lngReturnField = tOHDatabaseFields.OHUser1 Then
            GetOrderData = CStr(udt.OHUser1)
        End If
    End If
    
    CloseOHDatabase
    dbOH = 0
    
    Exit Function
End Function
Public Function UserPasswordCheck(strUser As String, strPassword As String) As Boolean
    Dim strTemp As String
    Dim I As Integer
    Dim udtUser As tUserRecord
    UserPasswordCheck = False
    
  OpenUserDatabase     ' ltas 8.29.2006
    
    If dbUser = 0 Then
        MsgBox "Unable to open user database"
        Exit Function
    End If
    
    PDBSetSortFields dbUser, 1
    PDBMoveFirst dbUser
    
    PDBFindRecordByField dbUser, 1, strUser
    
    If PDBGetLastError(dbUser) = ErrNone Then
        PDBReadRecord dbUser, VarPtr(udtUser)
        gstrUserName = strUser
        Do Until UCase(udtUser.UserLoginNameHH) <> UCase(gstrUserName)
            If Trim(UCase(udtUser.UserPasswordHH)) = Trim(UCase(strPassword)) Then
                UserPasswordCheck = True
                Exit Do
            Else
                PDBMoveNext dbUser
                PDBReadRecord dbUser, VarPtr(udtUser)
                If PDBEOF(dbUser) = True Then Exit Do
            End If
        Loop
        
    End If
    
    If UserPasswordCheck = True Then
        ReDim garySecurity(0)
        If SC(gstrUserName, "001") = True Then
                garySecurity(0) = "HHADMIN"
        Else
            If InStr(udtUser.UserSecurityHH, "~") > 0 Then
                Call AppForge_Split(udtUser.UserSecurityHH, pstrSplit, "~")
                For I = 0 To UBound(pstrSplit.ary)
                    If I > 0 Then ReDim Preserve garySecurity(I)
                    garySecurity(I) = pstrSplit.ary(I)
                Next
            Else
                garySecurity(0) = udtUser.UserSecurityHH
            End If
        End If
    End If
                
    PDBClose (dbUser)
    
    Exit Function
    
    
End Function

Function fgetFootage(intLength As Double, intWidth As Double, intPieces As Integer, strThickness As String, _
    Optional strChainTally As String = "", Optional dblThkCalcFactor As Double, Optional strRoundUpDown As String) As Double
    Dim sngThickness As Single
    Dim strBeginThickness As String
    Dim dblCalcFactor As Double
    
    
On Error GoTo ErrorHandler

    sngThickness = 1
    strBeginThickness = strThickness
    
    If SC(strChainTally, "BTMetric") = True Then
        Dim intCBMDivisor As Integer
        Dim dblFactor As Double
        
        If CDbl(gSettings.SystemM3_BFMConversionFactor) < 0.005 Then
            intCBMDivisor = 1
        Else
            intCBMDivisor = 12
        End If
        dblFactor = (CDbl(gSettings.SystemM3_BFMConversionFactor) / intCBMDivisor)
        
        
        intLength = Round(intLength / CDbl(gSettings.systemM3_FeetToCM_ConversionFactor), 3)
        intWidth = Round(intWidth / CDbl(gSettings.systemM3_InchesToCM_ConversionFactor), 3)
    End If
    
    
    If (dblThkCalcFactor) > 0 Then
        sngThickness = CSng(dblThkCalcFactor)
    Else
        'If there is a calcfactor for this thickness, then get it and we will use it for the calculation instead of all this below stuff...
        Call AppForge_Split(GetThicknessData(Trim(UCase(strThickness)), "Thickness"), pstrSplit, ",")
        
        If IsNumeric(pstrSplit.ary(3)) = True Then
            If CDbl(pstrSplit.ary(3)) > 0 Then
                dblCalcFactor = CDbl(pstrSplit.ary(3))
            
            ElseIf IsNumeric(strThickness) = True And SC(strChainTally, "BTMETRIC") = True Then
                dblCalcFactor = (CDbl(strThickness) / 10 / CDbl(gSettings.systemM3_InchesToCM_ConversionFactor))
            End If
            
        End If
        
        If dblCalcFactor > 0 Then
            sngThickness = dblCalcFactor
        Else
            
            If InStr(strThickness, "-") > 0 And InStr(strThickness, "/") > 0 Then
                'Get the Whole Number
                Call AppForge_Split(strThickness, pstrSplit, "-")
                sngThickness = CDbl(pstrSplit.ary(0))
                
                'Get the Partial Value which must be in the format #-#/#
                strThickness = pstrSplit.ary(1)
                Call AppForge_Split(strThickness, pstrSplit, "/")
                If UBound(pstrSplit.ary) = 1 Then
                    sngThickness = sngThickness + (CDbl(pstrSplit.ary(0)) / CDbl(pstrSplit.ary(1)))
                End If
                    
                strThickness = strBeginThickness
        '''''don't know what the hell this was suppose to do   sngThickness = CDbl(pstrSplit.ary(0) & ".5") / 4
        
            Else
                If InStr(strThickness, "/") > 0 Then
                    Call AppForge_Split(strThickness, pstrSplit, "/")
                ElseIf dblCalcFactor = 0 Then
                    Call AppForge_Split(strThickness & "/4", pstrSplit, "/")
                End If
                
                If UBound(pstrSplit.ary) = 1 Then
                    If IsNumeric(pstrSplit.ary(0)) = True And IsNumeric(pstrSplit.ary(1)) = True Then
                    
                        sngThickness = CDbl(pstrSplit.ary(0)) / CDbl(pstrSplit.ary(1))
                    Else
                        'not normally structured thickness with /4 in it.
                    End If
                End If
            End If
        End If
    End If

    If gFootageTrueMath = True Or strChainTally = "CHAINTALLY" Then
       fgetFootage = (intWidth / 12) * intLength * intPieces * sngThickness
    Else
        fgetFootage = cRound((Round(intWidth, 3) * intLength * sngThickness) / 12) * intPieces
    End If
    
    If SC(gSettings.ETUseRoundMath, "YES") = True Then
        If Right(CStr(fgetFootage), 2) = ".5" Then
            If SC(strRoundUpDown, "UP") = True Then
                fgetFootage = fgetFootage + 0.5
            Else
                fgetFootage = CDbl(Replace(CStr(fgetFootage), ".5", ".0"))
            End If
        Else
            fgetFootage = Round(fgetFootage, 0)
        End If
    End If
    
    Exit Function
ErrorHandler:
    fgetFootage = 0
    Exit Function
    
End Function

Public Function cRound(dblNumber As Double) As Double
    Dim dblTemp As Double
    dblTemp = dblNumber - CInt(dblNumber)
    
    If dblTemp < 0 Then
        cRound = CInt(dblNumber)
    ElseIf dblTemp = 0 Then
        cRound = dblNumber
    ElseIf dblTemp < 0.5 Then
        cRound = CInt(dblNumber)
    ElseIf dblTemp >= 0.5 Then
        cRound = CInt(dblNumber) + 1
    Else
        MsgBox ("Error Message in Subroutine cRound. Unexpected Conclusion.")
    End If
End Function

Public Sub getCulls()
Dim udt As tUtilityRecord
Dim lfound As Boolean
     OpenUtilityDatabase
     PDBMoveFirst (dbUtility)
     Do While PDBEOF(dbUtility) = False
        PDBReadRecord dbUtility, VarPtr(udt)
        If udt.UtilityName = "CULLS" Then
            gCullPieces = CDbl(udt.UtilityValueText)
            gCullFootage = CDbl(udt.UtilityValue2Text)
            lfound = True
            Exit Do
        End If
        PDBMoveNext (dbUtility)
     Loop
     CloseUtilityDatabase
     If lfound = False Then Call setCulls(0, 0)
End Sub

Public Sub setCulls(pcs As Double, ftg As Double)
Dim udt As tUtilityRecord
Dim lfound As Boolean
lfound = False
     OpenUtilityDatabase
     PDBMoveFirst (dbUtility)
     Do While PDBEOF(dbUtility) = False
        PDBReadRecord dbUtility, VarPtr(udt)
        If udt.UtilityName = "CULLS" Then
            PDBEditRecord (dbUtility)
            udt.UtilityValueText = CStr(gCullPieces)
            udt.UtilityValue2Text = CStr(gCullFootage)
            PDBWriteRecord dbUtility, VarPtr(udt)
            PDBUpdateRecord (dbUtility)
            lfound = True
            Exit Do
        End If
        PDBMoveNext (dbUtility)
     Loop
     
If lfound = False Then
    PDBCreateRecordBySchema (dbUtility)
    'udt.UtilityID = GetNextIDPDB(dbUtility, 0)
    udt.UtilityName = "CULLS"
    udt.UtilityValueText = CStr(gCullPieces)
    udt.UtilityValue2Text = CStr(gCullFootage)
    PDBWriteRecord dbUtility, VarPtr(udt)
    PDBUpdateRecord dbUtility
End If
     
     
     CloseUtilityDatabase
End Sub
Public Sub getPartnerMode()

    Dim udt As tUtilityRecord
    Dim lfound As Boolean
     
     OpenUtilityDatabase
     PDBMoveFirst (dbUtility)
     
     Do While PDBEOF(dbUtility) = False
        PDBReadRecord dbUtility, VarPtr(udt)
        If udt.UtilityName = "PARTNERMODE" Then
            gstrPartnerMode = CStr(udt.UtilityValueText)
            lfound = True
            Exit Do
        End If
        PDBMoveNext (dbUtility)
     Loop
     
     CloseUtilityDatabase
    dbUtility = 0

End Sub




Public Sub getDailyFootage()
Dim udt As tUtilityRecord
Dim lfound As Boolean
     OpenUtilityDatabase
     PDBMoveFirst (dbUtility)
     Do While PDBEOF(dbUtility) = False
        PDBReadRecord dbUtility, VarPtr(udt)
        If udt.UtilityName = "DAILYFOOTAGE" Then
            gDailyDay = CDbl(udt.UtilityValueText)
            gDailyFootage = CDbl(udt.UtilityValue2Text)
            lfound = True
            Exit Do
        End If
        PDBMoveNext (dbUtility)
     Loop
     CloseUtilityDatabase
     If lfound = False Then Call setCulls(0, 0)
End Sub

Public Sub setDailyFootage(day As Integer, ftg As Double)
Dim udt As tUtilityRecord
Dim lfound As Boolean
lfound = False
     OpenUtilityDatabase
     PDBMoveFirst (dbUtility)
     Do While PDBEOF(dbUtility) = False
        PDBReadRecord dbUtility, VarPtr(udt)
        If udt.UtilityName = "DAILYFOOTAGE" Then
            PDBEditRecord (dbUtility)
            udt.UtilityValueText = CStr(gDailyDay)
            udt.UtilityValue2Text = CStr(gDailyFootage)
            PDBWriteRecord dbUtility, VarPtr(udt)
            PDBUpdateRecord (dbUtility)
            lfound = True
            Exit Do
        End If
        PDBMoveNext (dbUtility)
     Loop
     
If lfound = False Then
    PDBCreateRecordBySchema (dbUtility)
    'udt.UtilityID = GetNextIDPDB(dbUtility, 0)
    udt.UtilityName = "DAILYFOOTAGE"
    udt.UtilityValueText = CStr(gDailyDay)
    udt.UtilityValue2Text = CStr(gDailyFootage)
    PDBWriteRecord dbUtility, VarPtr(udt)
    PDBUpdateRecord dbUtility
End If
          
     CloseUtilityDatabase
End Sub



Public Sub LoadgudtPD(strBundle As String, ByRef udtPD As tPDRecord, strBundleSearchType As String, lngBundleID As Long)
       
On Error GoTo ErrorHandler
    OpenPDDatabase
        
    
    If strBundleSearchType = "PDID" Then
        PDBSetSortFields dbPD, 0
        PDBMoveFirst dbPD
        
        PDBFindRecordByField dbPD, 0, lngBundleID
    Else
        PDBSetSortFields dbPD, 1
        PDBMoveFirst dbPD
        
        PDBFindRecordByField dbPD, 1, UCase(strBundle)
    End If
    
    If PDBGetLastError(dbPD) = ErrNone Then
        PDBReadRecord dbPD, VarPtr(gudtPD)
    Else
        gudtPD.BundleID = ""
        MsgBox "Production Bundle: " & strBundle & " Not Found!"
        Exit Sub
    End If
    
    PDBClose dbPD
    Exit Sub
ErrorHandler:
    MsgBox "Error in LoadgudtPD " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub
Public Sub Utility_Search(strUtiltyName As String, strUtilityValueText As String, strUtilitySearchType As String, _
fSingleResult As Boolean, Optional fAutoAddIfNotExists_NameSingleValOnly As Boolean, Optional fAutoUpdatewithIncomingValue As Boolean, Optional strFieldtoUpdate As String)
    Dim fInsert As Boolean

On Error GoTo ErrorHandler

            
    If dbUtility = 0 Then OpenUtilityDatabase
    ReDim gudtUtility(0)
    
    If dbUtility = 0 Then
        MsgBox "Unable to open utility database"
        Exit Sub
    End If
        
    
    If strUtilitySearchType = "NAME/VALUETEXT" And fSingleResult = True Then
        PDBSetSortFields dbUtility, tUtilityDatabaseFields.UtilityName_Field
        PDBMoveFirst dbUtility
        
        PDBFindRecordByField dbUtility, 1, strUtiltyName
        ReDim gudtUtility(0)
        Do Until PDBEOF(dbUtility)
            PDBReadRecord dbUtility, VarPtr(gudtUtility(0))
            If SC(gudtUtility(0).UtilityName, strUtiltyName) = True And SC(gudtUtility(0).UtilityValueText, strUtilityValueText) Then
                'we are done here
                Exit Do
            Else
                ReDim gudtUtility(0)
            End If
            PDBMoveNext dbUtility
        Loop
    
    ElseIf strUtilitySearchType = "NAME/UTILITYVALUELONG" And fSingleResult = True Then
        PDBSetSort dbUtility, "UtilityName,UtilityValueLong"
        PDBMoveFirst dbUtility
        
        PDBFindRecordByField dbUtility, 1, strUtiltyName
        ReDim gudtUtility(0)
        Do Until PDBEOF(dbUtility)
            PDBReadRecord dbUtility, VarPtr(gudtUtility(0))
            If SC(gudtUtility(0).UtilityName, strUtiltyName) And SC(CStr(gudtUtility(0).UtilityValueLong), CLng(strUtilityValueText)) Then
                'we are done here
                Exit Do
            Else
                ReDim gudtUtility(0)
            End If
            PDBMoveNext dbUtility
        Loop
    ElseIf strUtilitySearchType = "NAME" And fSingleResult = True Then
        
        PDBSetSortFields dbUtility, 1
        PDBMoveFirst dbUtility
        
        PDBFindRecordByField dbUtility, 1, strUtiltyName
        ReDim gudtUtility(0)
        Do Until PDBEOF(dbUtility)
            PDBReadRecord dbUtility, VarPtr(gudtUtility(0))
            If SC(gudtUtility(0).UtilityName, strUtiltyName) Then
                If fAutoUpdatewithIncomingValue = True Then
                    If strFieldtoUpdate = "" Or SC(strFieldtoUpdate, "UtilityValueText") Then
                        gudtUtility(0).UtilityValueText = strUtilityValueText
                    ElseIf strFieldtoUpdate = "UtilityValueLong" Then
                        If IsNumeric(strUtilityValueText) = True Then
                            gudtUtility(0).UtilityValueLong = CLng(strUtilityValueText)
                        Else
                            gudtUtility(0).UtilityValueLong = 0
                        End If
                    End If
                    gudtUtility(0).UtilityValueText = strUtilityValueText
                    PDBEditRecord dbUtility
                    fInsert = WriteUtilityRecord(gudtUtility(0))
                    PDBUpdateRecord dbUtility
                End If
                
                'we are done here
                'MsgBox "found Update record"
                Exit Do
            Else
                ReDim gudtUtility(0)
            End If
            PDBMoveNext dbUtility
        Loop
        
        If SC(gudtUtility(0).UtilityName, strUtiltyName) = False And fAutoAddIfNotExists_NameSingleValOnly = True Then
            gudtUtility(0).UtilityName = strUtiltyName
            gudtUtility(0).UtilityID = PDBCreateDBUniqueNumber(dbUtility)
            gudtUtility(0).UtilityValue2Long = 0
            gudtUtility(0).UtilityValue2Text = ""
            gudtUtility(0).UtilityValueLong = 0
            gudtUtility(0).UtilityValueText = ""
            
            'Initial Default value set for new variable
            If SC(strUtiltyName, "UNITDEFAULT") Then
                gudtUtility(0).UtilityValueText = "USA"
            ElseIf SC(strUtiltyName, "ALLOWNONINVBUNDLEADD") Then
                gudtUtility(0).UtilityValueText = "NO"
            ElseIf SC(strUtiltyName, "PRODFIELD") Then
                gudtUtility(0).UtilityValueText = "PRODDESC"
            ElseIf SC(strUtiltyName, "XFERSEEDVALUE") Then
                gudtUtility(0).UtilityValueText = "100"
            ElseIf SC(strUtiltyName, "XFERSEEDPREFIX") Then
                gudtUtility(0).UtilityValueText = "T"
            ElseIf SC(strUtiltyName, "SHIPSEEDVALUE") Then
                gudtUtility(0).UtilityValueText = "100"
            ElseIf SC(strUtiltyName, "SHIPSEEDPREFIX") Then
                gudtUtility(0).UtilityValueText = "S"
            ElseIf SC(strUtiltyName, "INVSEEDVALUE") Then
                gudtUtility(0).UtilityValueText = "100"
            ElseIf SC(strUtiltyName, "INVSEEDPREFIX") Then
                gudtUtility(0).UtilityValueText = "I"
            ElseIf SC(strUtiltyName, "KILNSEEDVALUE") Then
                gudtUtility(0).UtilityValueText = "100"
            ElseIf SC(strUtiltyName, "KILNSEEDPREFIX") Then
                gudtUtility(0).UtilityValueText = "K"
            ElseIf SC(strUtiltyName, "XFERALLOWNOTFOUND") Then
                gudtUtility(0).UtilityValueText = "NO"
            ElseIf SC(strUtiltyName, "CLIENT") Then
                gudtUtility(0).UtilityValueText = "CLIENT"
            ElseIf SC(strUtiltyName, "SCANREMOVELEADING") Then
                gudtUtility(0).UtilityValueText = "NO"
            ElseIf SC(strUtiltyName, "SEEDVALUEET") Then
                gudtUtility(0).UtilityValueText = "100000"
            ElseIf SC(strUtiltyName, "SEEDPREFIXET") Then
                gudtUtility(0).UtilityValueText = "U"
            ElseIf SC(strUtiltyName, "SEEDVALUEETUSE") Then
                gudtUtility(0).UtilityValueText = "0"
            ElseIf SC(strUtiltyName, "ETREQ") Then
                gudtUtility(0).UtilityValueText = ""
            ElseIf SC(strUtiltyName, "RECEIVEESTIMATEMETHOD") Then
                gudtUtility(0).UtilityValueText = "LENGTH"
            ElseIf SC(strUtiltyName, "BTCOMMRECEIVE") Then
                gudtUtility(0).UtilityValueLong = gintPrinterPort
                gudtUtility(0).UtilityValueText = CStr(gintPrinterPort)

            ElseIf SC(strUtiltyName, "TAGFILE-ET") Then
                gudtUtility(0).UtilityValueText = "ENDTALLY.LBL"
            ElseIf SC(strUtiltyName, "SHIPOVERAGEPERCENT") Then
                gudtUtility(0).UtilityValueText = "0.10"
            ElseIf SC(strUtiltyName, "RECAIDREQUIRED") Then
                gudtUtility(0).UtilityValueText = "0"
            ElseIf SC(strUtiltyName, "ENDTALLYMAXWIDTH") Then
                gudtUtility(0).UtilityValueText = "29"
            ElseIf SC(strUtiltyName, "ENDTALLYTWOKEYMAX") Then
                gudtUtility(0).UtilityValueText = "2"
            ElseIf SC(strUtiltyName, "ETKEYF3LOCATION") Or SC(strUtiltyName, "ETKEYF4LOCATION") Or SC(strUtiltyName, "ETKEYF5LOCATION") _
                Or SC(strUtiltyName, "ETKEYF6LOCATION") Then
                gudtUtility(0).UtilityValueText = ""
            ElseIf SC(strUtiltyName, "RUNIDLOOKUPTYPE") Then
                gudtUtility(0).UtilityValueText = ""
            ElseIf SC(strUtiltyName, "BACKUPTYPE") Then
                gudtUtility(0).UtilityValueText = "STANDARD"
            ElseIf SC(strUtiltyName, "BACKUPCOUNT") Then
                gudtUtility(0).UtilityValueText = "1"
            ElseIf SC(strUtiltyName, "ETPRINTTYPE") Then
                gudtUtility(0).UtilityValueText = "MOBILE"
                
            ElseIf SC(strUtiltyName, "RECPRINTTYPE") Then
                gudtUtility(0).UtilityValueText = "MOBILE"
            ElseIf SC(strUtiltyName, "RECFILEPATH") Then
                gudtUtility(0).UtilityValueText = ""
            
            ElseIf SC(strUtiltyName, "RECPO") Then
                gudtUtility(0).UtilityValueText = ""
            ElseIf SC(strUtiltyName, "RECVENDOR") Then
                gudtUtility(0).UtilityValueText = ""
            ElseIf SC(strUtiltyName, "DEBUGMODEON") Then
                gudtUtility(0).UtilityValueText = "NO"
                
            ElseIf SC(strUtiltyName, "DEFAULTTHICKNESS") Then
                gudtUtility(0).UtilityValueText = "4"

            ElseIf SC(strUtiltyName, "LASTRUNID") Then
                gudtUtility(0).UtilityValueText = ""
            ElseIf SC(strUtiltyName, "DefaultDimensioned") Then
                gudtUtility(0).UtilityValueText = "0"
            ElseIf SC(strUtiltyName, "FRMLOOKUPMAXWIDTHVALUE") Then
                gudtUtility(0).UtilityValueText = "1000"
            ElseIf SC(strUtiltyName, "WIPSEEDVALUE") Then
                gudtUtility(0).UtilityValueText = "100"
            ElseIf SC(strUtiltyName, "KILNFORMVERSION") Then
                gudtUtility(0).UtilityValueText = "1"
            ElseIf SC(strUtiltyName, "ORDERALLOWOVERRIDE") Then
                gudtUtility(0).UtilityValueText = "NO"
            ElseIf SC(strUtiltyName, "BUNDLETODAYETFORM") Then
                gudtUtility(0).UtilityValueText = "FRMBUNDLEHEADER"
            ElseIf SC(strUtiltyName, "BTCLASSREQ") Then
                gudtUtility(0).UtilityValueText = "YES"
            ElseIf SC(strUtiltyName, "BTWIDTHREQ") Then
                gudtUtility(0).UtilityValueText = "YES"
            ElseIf SC(strUtiltyName, "TOPOFFORMLOCK") Then
                gudtUtility(0).UtilityValueText = "25"
            
            Else
                gudtUtility(0).UtilityValueText = strUtilityValueText
                If SC(gudtUtility(0).UtilityName, "IPSM") Then
On Error GoTo SkipError
'PDINV                    Call CreateDatabasePDB(dbPDINV, "PDINV", PDINV_Schema)
'PDINV                    Dim fso As New CFileManager
                    
'PDINV                    fso.MoveFile App.Path & "\PDINV.pdb", gstrIPSM & "\PDINV.pdb"
'PDINV                    Set fso = Nothing
SkipError:
                End If
                
                
            End If
                        
            'Make sure we don't get a duplicated ID for the pdb table, happens when using the pdbcreateuni....call sometimes
            Dim fUtilityIDDupCheck As Boolean
            fUtilityIDDupCheck = True
            Do Until fUtilityIDDupCheck = False
                Dim udt As tUtilityRecord
                PDBMoveFirst dbUtility
                PDBFindRecordByField dbUtility, tUtilityDatabaseFields.UtilityID_Field, VarPtr(udt)
                
                fUtilityIDDupCheck = False 'reset to false then check see if there one exists already
                If PDBGetLastError(dbUtility) = 0 Then
                    'record was found, so it's dup'd
                    fUtilityIDDupCheck = True
                End If
            Loop
            
            PDBCreateRecordBySchema (dbUtility)
            PDBWriteRecord dbUtility, VarPtr(gudtUtility(0))
            PDBUpdateRecord dbUtility
        End If
        
    End If
    
    PDBClose dbUtility
    dbUtility = 0
    
    Exit Sub

    Exit Sub
ErrorHandler:
    MsgBox "Error in Utility_Search " & Err.Number & "-" & Err.Description
    Exit Sub
End Sub
Public Sub Utility_Delete(strUtiltyName As String, fDeleteAll_THISDOESNOTHING As Boolean, strUpdateName As String)
       
    Call Utility_Search("SHRNKREPAIR", "", "NAME", True)
    
    If gudtUtility(0).UtilityName = "SHRNKREPAIR" Then
        'the udpate has already been run
        Exit Sub
    End If
    CloseUtilityDatabase
    OpenUtilityDatabase
    ReDim gudtUtility(0)
    
    If dbUtility = 0 Then
        MsgBox "Unable to open utility database"
        Exit Sub
    End If
        
    
    PDBSetSortFields dbUtility, tUtilityDatabaseFields.UtilityName_Field
    PDBMoveFirst dbUtility
    
    PDBFindRecordByField dbUtility, 1, strUtiltyName
    ReDim gudtUtility(0)
    
    Do Until PDBEOF(dbUtility)
        PDBReadRecord dbUtility, VarPtr(gudtUtility(0))
        If SC(gudtUtility(0).UtilityName, strUtiltyName) Then
            PDBDeleteRecordEx dbUtility, afDeleteModeRemove
        Else
            ReDim gudtUtility(0)
            Exit Do
        End If
    Loop
    
    If strUpdateName <> "" Then
        gudtUtility(0).UtilityID = PDBCreateDBUniqueNumber(dbUtility)
        gudtUtility(0).UtilityName = strUpdateName
        gudtUtility(0).UtilityValue2Long = 0
        gudtUtility(0).UtilityValue2Text = ""
        gudtUtility(0).UtilityValueLong = 0
        gudtUtility(0).UtilityValueText = ""
        
        PDBCreateRecordBySchema dbUtility
        PDBWriteRecord dbUtility, VarPtr(gudtUtility(0))
        PDBUpdateRecord dbUtility
    End If
    
    PDBClose dbUtility
    dbUtility = 0
    
    Exit Sub
    
End Sub

Public Sub ErrorCheck(db As Long, Optional strLastAction As String, Optional fRecordSearch As Boolean)

    If PDBGetLastError(db) <> ErrNone Then
        If PDBGetLastError(db) = -7 And fRecordSearch = True Then
            'do nothing...don't send any error back...just couldn't find it
        ElseIf PDBGetLastError(db) <> -33 Then
            MsgBox "DB Error " & PDBGetLastError(db) & " _ LastAction:" & strLastAction
        End If
    End If
        
End Sub
Public Sub LoadgaryCTLine()
    Dim I As Long
    
    ReDim garyCTLine(0)
    garyCTLine(0).CTLineID = -1
    
    '8.24.2006 ltas dbCTLine = PDBOpen(Byfilename, gstrPDBPath & "\CTLine", 0, 0, 0, 0, afModeReadWrite)
    OpenCTLineDatabase
    Call ErrorCheck(dbCTLine)
    If dbCTLine = 0 Then
        MsgBox "Unable to open Chain Tally database"
        Exit Sub
    End If
    
    PDBSetSort dbCTLine, "LoadAID,Species,Thickness,Grade"
    
    PDBMoveFirst dbCTLine
    I = -1
    Do Until PDBEOF(dbCTLine)
        I = I + 1
        If I = 0 Then
            ReDim garyCTLine(0)
            garyCTLine(0).CTLineID = -1
        Else
            ReDim Preserve garyCTLine(I)
        End If
        
        Call ReadCTLineRecord(garyCTLine(I))
        PDBMoveNext dbCTLine
    Loop
            
   CloseCTLineDatabase
    
End Sub
Public Sub LoadgaryLR()
    Dim I As Long
    
    ReDim garyLR(0)
    garyLR(0).LRID = -1
    
    '8.24.2006 ltas dbLR = PDBOpen(Byfilename, gstrPDBPath & "\LR", 0, 0, 0, 0, afModeReadWrite)
    OpenLRDatabase
    Call ErrorCheck(dbLR)
    If dbLR = 0 Then
        MsgBox "Unable to open Load Receipt Database"
        Exit Sub
    End If
    
    PDBSetSortFields dbLR, 1
    
    PDBMoveFirst dbLR
    I = -1
    Do Until PDBEOF(dbLR)
        I = I + 1
        If I = 0 Then
            ReDim garyLR(0)
            garyLR(0).LRID = -1
        Else
            ReDim Preserve garyLR(I)
        End If
        
        Call ReadLRRecord(garyLR(I))
        PDBMoveNext dbLR
    Loop
            
   CloseLRDatabase
    
End Sub

Public Sub LoadgaryCTLineLast2()
    Dim I As Integer
    
     ''8.24.2006   dbCTLine = PDBOpen(Byfilename, gstrPDBPath & "\CTLine", 0, 0, 0, 0, afModeReadWrite)
     
     OpenCTLineDatabase
     
    Call ErrorCheck(dbCTLine)
    If dbCTLine = 0 Then
        MsgBox "Unable to open Chain Tally database"
        Exit Sub
    End If
    PDBSetSortFields dbCTLine, 0
    PDBMoveLast dbCTLine
    If PDBNumRecords(dbCTLine) > 1 Then
        PDBMovePrev dbCTLine
    End If
    
    I = -1
    Do Until PDBEOF(dbCTLine)
        I = I + 1
        If I = 0 Then
            ReDim garyCTLine(0)
            garyCTLine(0).CTLineID = -1
        Else
            ReDim Preserve garyCTLine(I)
        End If
        
        Call ReadCTLineRecord(garyCTLine(I))
        
        PDBMoveNext dbCTLine
    Loop
    
    
    PDBFindRecordByField dbCTLine, 0, garyCTLine(UBound(garyCTLine)).CTLineID
    
    
    PDBDeleteRecordEx dbCTLine, afDeleteModeRemove
    
    
    
    PDBGetLastError (dbCTLine)
    
    
    CloseCTLineDatabase
    
End Sub

Public Sub LoadgaryCTLineLast2_NEW()
'''    Dim I As Integer
'''
'''     dbCTLine = PDBOpen(Byfilename, gstrPDBPath & "\CTLine", 0, 0, 0, 0, afModeReadWrite)
'''    Call ErrorCheck(dbCTLine)
'''    If dbCTLine = 0 Then
'''        MsgBox "Unable to open Chain Tally database"
'''        Exit Sub
'''    End If
'''    PDBSetSortFields dbCTLine, 0
'''    PDBMoveLast dbCTLine
'''    If PDBNumRecords(dbCTLine) > 1 Then
'''        PDBMovePrev dbCTLine
'''    End If
'''
'''    I = -1
'''    Do Until PDBEOF(dbCTLine)
'''        I = I + 1
'''        If I = 0 Then
'''            ReDim garyCTLine(0)
'''            garyCTLine(0).CTLineID = -1
'''        Else
'''            ReDim Preserve garyCTLine(I)
'''        End If
'''
'''        Call ReadCTLineRecord(garyCTLine(I))
'''
'''        PDBMoveNext dbCTLine
'''    Loop
'''
'''
'''    'Modify this section Leith.Stetson 7.13.2006
'''
'''    PDBFindRecordByField dbCTLine, 28, GetCTLineAutoMatch()
'''
'''
'''    PDBDeleteRecordEx dbCTLine, afDeleteModeRemove
'''
'''
'''
'''    PDBGetLastError (dbCTLine)
'''
'''
'''    PDBClose dbCTLine
'''    dbCTLine = 0
    
End Sub


Public Sub LoadgaryPD(strid As String, strSearchType As String)
    
    Dim fmatch As Boolean
    
    Dim udtPD As tPDRecord
    Dim I As Integer
    
    OpenPDDatabase
    
    If strSearchType = "BOL" Then
        PDBSetSortFields dbPD, 15
        PDBMoveFirst dbPD
        PDBFindRecordByField dbPD, 15, CLng(strid)
    ElseIf strSearchType = "KILN" Then
        PDBSetSort dbPD, "KilnNumber,BundleID"
        PDBFindRecordByField dbPD, 30, CInt(strid)
    Else
        PDBSetSortFields dbPD, 0
    End If
            
    If PDBGetLastError(dbPD) = ErrNone Then
        fmatch = True
        ReDim garyPD(0)
        I = 0
        Do Until fmatch = False
            
            'Add one to the array of results and then read the next record into it
                'if it shouldn't be read, it will be removed later
            If garyPD(0).PDID <= 0 Then
                'Do Nothing...no need to redim it is already empty
            Else
                I = I + 1
                ReDim Preserve garyPD(I)
            End If
            
            PDBReadRecord dbPD, VarPtr(garyPD(I))
            
            If strSearchType = "BOL" Then
                If garyPD(I).BOLID = CLng(strid) Then
                    fmatch = True
                Else
                    If I = 0 Then
                        ReDim garyPD(0)
                        Exit Sub
                    Else
                        fmatch = False
                    ReDim Preserve garyPD(I - 1)
                    End If
                End If
            ElseIf strSearchType = "KILN" Then
                If garyPD(I).KilnNumber = CInt(strid) And garyPD(I).StatusID = STATUS_KILN Then
                    fmatch = True
                ElseIf garyPD(I).KilnNumber = CInt(strid) Then
                    'Right Kiln...wrong status...keep looking :)
                    If I = 0 Then
                        If garyPD(I).KilnNumber = CInt(strid) And garyPD(I).StatusID = STATUS_KILN Then
                            'Do nothing
                            'MsgBox Now
                        Else    'remove it
                        ReDim garyPD(0)
                        End If
                        fmatch = True
                    Else
                        fmatch = True
                        ReDim Preserve garyPD(I - 1)
                        I = I - 1
                    End If
                Else
                    If I = 0 Then
                        ReDim garyPD(0)
                        Exit Sub
                    Else
                        fmatch = False
                        ReDim Preserve garyPD(I - 1)
                    End If
                End If
            ElseIf strSearchType = "ALL" Then
                fmatch = True
            End If
            
            PDBMoveNext dbPD
            If PDBEOF(dbPD) = True Then fmatch = False
        Loop
                        
    Else
        ReDim garyPD(0)
    End If
    
    PDBSetSortFields dbPD, 1
    ClosePDDatabase
    
End Sub

Public Sub LoadgudtPDLine(lngPDID As Long, Optional strSortString As String, Optional ShiftID As String = "", Optional keepprevious As Boolean = False)
    
    Dim fmatch As Boolean
    Dim udtPDLine As tPdLineRecord
    Dim I As Integer
    Dim lngPieces As Integer
    
On Error GoTo ErrorHandler

    If keepprevious = False Then
        ReDim gudtPDLine(0)
    End If
    
    OpenPdLineDatabase     ' ltas 8.29.2006
    
    If dbPDLine = 0 Then
        MsgBox "Unable to open production line database"
        Exit Sub
    End If
    
    If strSortString = "" Then
        PDBSetSortFields dbPDLine, tPdLineDatabaseFields.PDID_Field
    Else
        PDBSetSort dbPDLine, strSortString
    End If

    
    PDBMoveFirst dbPDLine
     
    PDBFindRecordByField dbPDLine, 1, lngPDID
    'MsgBox lngPDID
        
    If PDBGetLastError(dbPDLine) = ErrNone Then
        fmatch = True
        
        I = UBound(gudtPDLine)
        Do Until fmatch = False
            PDBReadRecord dbPDLine, VarPtr(udtPDLine)
'            Debug.Print udtPDLine.PDLineID
            If udtPDLine.PDID = lngPDID Then
                If gudtPDLine(0).PDLineID = 0 Then
                    'Do Nothing...no need to redim it is already empty
                Else
                    I = I + 1
                    ReDim Preserve gudtPDLine(I)
                End If
                
                gudtPDLine(I).PDLineID = udtPDLine.PDLineID
                gudtPDLine(I).PDID = udtPDLine.PDID
                gudtPDLine(I).GradeID = udtPDLine.GradeID
                gudtPDLine(I).PDLineLength = udtPDLine.PDLineLength
                gudtPDLine(I).PDLineWidth = udtPDLine.PDLineWidth
                gudtPDLine(I).PDLinePieces = udtPDLine.PDLinePieces
                lngPieces = lngPieces + udtPDLine.PDLinePieces
                
                gudtPDLine(I).PDLineGross = udtPDLine.PDLineGross
                gudtPDLine(I).PDLineNet = udtPDLine.PDLineNet
                gudtPDLine(I).PDLineNote = udtPDLine.PDLineNote
                gudtPDLine(I).GID = udtPDLine.GID
                gudtPDLine(I).GDT = udtPDLine.GDT
                gudtPDLine(I).thickness = udtPDLine.thickness
                gudtPDLine(I).ThicknessID = udtPDLine.ThicknessID
                gudtPDLine(I).Other = udtPDLine.Other
                gudtPDLine(I).ShiftID = udtPDLine.ShiftID
                gudtPDLine(I).MillID = udtPDLine.MillID
            Else
                fmatch = False
            End If
         '   Debug.Print gudtPDLine(I).PDLineNet & "-" & gudtPDLine(I).PDLineGross
            
            PDBMoveNext dbPDLine
            If PDBEOF(dbPDLine) = True Then fmatch = False
        Loop
        
                        
    Else
        
        Exit Sub
    End If
    
    PDBClose dbPDLine
    
Exit Sub
ErrorHandler:
    MsgBox "Error in LoadGudtPDLine: " & Err.Number & "-" & Err.Description
    Exit Sub
End Sub

Public Function Bundle_RecalculateFootage_Any(strBundleID As String, lngPDID As Long, _
            ByVal fSaveValuesToPDTable As Boolean, _
            ByRef strReturnResults As String, ByRef strReturnTallyDetail As String, _
            ByRef aryReturnGrades As aryStringType, Optional ByRef dblReturnNetTotal As Double, _
            Optional ByRef dblReturnGrossTotal As Double, Optional dblReturnPcsTotal As Double) As Boolean
            
    Dim dblNetStart As Double, dblGrossStart As Double
    Dim dblNetEnd As Double, dblGrossEnd As Double
    Dim dblPcsStart As Double, dblPcsEnd As Double
    
    Dim strTallyDetailStart As String
    Dim strTallyorChain As String
    
    Dim fisGreen As Boolean
    Dim I As Long
    Dim int12 As Integer
    
On Error GoTo ErrorHandler
    ReDim gudtPDLine(0)
    dblNetEnd = 0
    dblGrossEnd = 0
    strTallyorChain = ""
    
    'Load the Bundle Header to global pd/pdline variables
    If lngPDID > 0 Then
        Call LoadgudtPD(strBundleID, gudtPD, "PDID", lngPDID)
    Else
        Call LoadgudtPD(strBundleID, gudtPD, "BUNDLEID", lngPDID)
    End If
    
    'Get Totals Pre-Recalc to Compare to values after recalc
    dblNetStart = Round(gudtPD.PDTotalNetBFM, 0)
    dblGrossStart = Round(gudtPD.PDTotalGrossBFM, 0)
    dblPcsStart = gudtPD.PDTotalPieces
    
    strTallyDetailStart = gudtPD.TallyDetail
    int12 = 1
    'Check Tally Type and make sure it's a mathematical calculation type (chain/end/dimensioned)
    If SC(gudtPD.TallyType, "LWCHAINTALLY") = True Then
        strTallyorChain = "CHAINTALLY"
        int12 = 1
    ElseIf SC(gudtPD.TallyType, "SMCHAINTALLY") = True Then
        strTallyorChain = "CHAINTALLY"
        int12 = 12
    ElseIf SC(gudtPD.TallyType, "BUNDLETALLY") = True Then
        strTallyorChain = "BUNDLETALLY"
    ElseIf InStr(tcu(gudtPD.TallyType), "DIMENSIONED") > 0 Then
        strTallyorChain = "DIMENSIONED"
    Else
        'It's an estimated bundle, therefore cannot be recalculated with this routine
        strTallyorChain = gudtPD.TallyType
        strReturnResults = "ERROR-Tally Type Does Not Allow Recalculation/No Piece Data"
        strReturnTallyDetail = "INVALID Bundle Type for Recalculation/Option"
        Bundle_RecalculateFootage_Any = True
        Exit Function
    End If
    
    
    
    'Load the PDLine Records to gudtPDLine()
    Call LoadgudtPDLine(gudtPD.PDID, "PDID,GradeID,PDLineLength,PDLineWidth", "", False)
    fisGreen = IsGreen(gudtPD.StatusID)
    
    Dim dblThkCalcFactor As Double
    Dim strLastThk As String
    
   'loop through line records and recalculate and round to two digits
    For I = 0 To UBound(gudtPDLine)
        If SC(strLastThk, gudtPDLine(I).thickness) = True And dblThkCalcFactor > 0 Then
            'no need to get calc factor ..already have it.
        Else
            dblThkCalcFactor = CDbl(GetThicknessData(gudtPDLine(I).thickness, "THICKNESS", "CALCFACTOR"))
            strLastThk = gudtPDLine(I).thickness
        End If
        
        'The int12 variable is a calculation correction factor used for Surface Measure Chain Tallies
        If fisGreen = True Then
            gudtPDLine(I).PDLineGross = Round(fgetFootage(gudtPDLine(I).PDLineLength, gudtPDLine(I).PDLineWidth, _
                        gudtPDLine(I).PDLinePieces, gudtPDLine(I).thickness, strTallyorChain, dblThkCalcFactor, "") * int12, 2)
            gudtPDLine(I).PDLineNet = Round(dblshrinkage * gudtPDLine(I).PDLineGross, 2)
        Else
            gudtPDLine(I).PDLineNet = Round(fgetFootage(gudtPDLine(I).PDLineLength, gudtPDLine(I).PDLineWidth, _
                        gudtPDLine(I).PDLinePieces, gudtPDLine(I).thickness, strTallyorChain, dblThkCalcFactor, "") * int12, 2)
            gudtPDLine(I).PDLineGross = Round(gudtPDLine(I).PDLineNet / dblshrinkage, 2)
        End If
        'Update Totals for Pieces/Gross/Net
        dblPcsEnd = dblPcsEnd + gudtPDLine(I).PDLinePieces
        dblGrossEnd = dblGrossEnd + gudtPDLine(I).PDLineGross
        dblNetEnd = dblNetEnd + gudtPDLine(I).PDLineNet
        
    Next
    
    Dim dblNetDiff As Double, dblGrossDiff As Double, dblPercentDiff As Double, dblPcsDiff As Double
    
    dblNetDiff = dblNetStart - Round(dblNetEnd, 0)
    dblGrossDiff = dblGrossStart - Round(dblGrossEnd, 0)
    dblPcsDiff = dblPcsStart - dblPcsEnd
    
    strReturnResults = ""
    
    If fisGreen = True Then
        dblPercentDiff = Round(dblGrossDiff / dblGrossStart * 100, 1)
        If dblGrossDiff = 0 Then
            strReturnResults = "Gross Ftg= No Change During Recalc"
        Else
            strReturnResults = "Gross Ftg Start/End= " & CStr(dblGrossStart) & " / " & CStr(dblGrossEnd) & " !!**CHANGED**!!"
        End If
    Else
        dblPercentDiff = Round(dblNetDiff / dblNetStart * 100, 1)
        If dblNetDiff = 0 Then
            strReturnResults = "Net Ftg-No Change During Recalc"
        Else
            strReturnResults = "Net Ftg Start/End= " & CStr(dblNetStart) & " / " & CStr(dblNetEnd) & " !!**CHANGED**!!"
        End If
    End If
    
    If dblPcsDiff = 0 Then
        strReturnResults = strReturnResults & vbCrLf & "Pieces-No Change During Recalc"
    Else
        strReturnResults = strReturnResults & vbCrLf & "Pieces Start/End= " & CStr(dblPcsStart) & " / " & CStr(dblPcsEnd) & " !!**CHANGED**!!"
    End If
    
    dblReturnPcsTotal = dblPcsEnd
    dblReturnGrossTotal = dblGrossEnd
    dblReturnNetTotal = dblNetEnd
    
    'Get the new Grade Breakdown and New Grade Array Sorted by Grade Sort Order with Footages by Grade
    If Bundle_GetMultiGradeTotals_GlobalPDVariables(fisGreen, aryReturnGrades, strReturnTallyDetail) = False Then
        strReturnResults = "Error Getting Grade Breakdown"
    End If
    
    If SC(strTallyDetailStart, strReturnTallyDetail) = True Then 'it's the same
        strReturnResults = strReturnResults & vbCrLf & _
                        "Tally Detail - No Change During Recalc"
    Else
        strReturnResults = strReturnResults & vbCrLf & vbCrLf & _
                        "Tally Detail Pre-Recalc=" & strTallyDetailStart & vbCrLf & _
                        "Tally Detal Post-Recalc=" & strReturnTallyDetail & " !!**CHANGED**!!"
    End If
    
    If fSaveValuesToPDTable = True Then
        gudtPD.TallyDetail = strReturnTallyDetail
        gudtPD.ActualGradeFootage = strReturnTallyDetail
        gudtPD.PDTotalGrossBFM = dblGrossEnd
        gudtPD.PDTotalNetBFM = dblNetEnd
        gudtPD.PDTotalPieces = dblPcsEnd
        
        Call SavegudtPDLineRecords(False, gudtPD.PDID, False)
        Call SavegudtPDRecord_ReturnMod5_TrueFalse(True, False)
        
    Else
        'Don't update bundle header fields and don't save the bundle
    End If
    Bundle_RecalculateFootage_Any = True
    
    Exit Function
ErrorHandler:
    MsgBox "Error in Bundle_RecalculateFootage_Any_GlobalPD() " & Err.Number & "-" & Err.Description
    Exit Function
End Function
Public Function Bundle_GetMultiGradeTotals_GlobalPDVariables(ByVal fisGreen As Boolean, _
                ByRef aryReturnGrades As aryStringType, ByRef strReturnTallyDetail As String) As Boolean
            
    Dim aryGrades() As String
    Dim fFound As Boolean
    Dim I As Long, J As Long
    Dim strGradeCurrent As String
    
On Error GoTo ErrorHandler
    
    ReDim aryGrades(3, 0) As String ' 0=GradeID, 1=HHAID, 2=Gross, 3=Net
    
    fFound = False
    
    
    
    fFound = False
    For I = 0 To UBound(gudtPDLine)
        'reset found varaible to false and check to see if grade is already in the grade array
        fFound = False
        For J = 0 To UBound(aryGrades, 2)
            If SC(aryGrades(0, J), CStr(gudtPDLine(I).GradeID)) = True Then
                'already in array, add the footage (gross/net) to the totals below
                fFound = True
                Exit For ' exit to outer loop/add totals or add grade and initialize totals
            Else
                fFound = False
                'continue check other positions in grade array, then below it will add it if not found
            End If
        Next
        
        'if ffound=true then just add line footage to grade totals, if =false then new position in array and initalize footages to zero, then add
        If fFound = True Then
            'just add to toals, already in the grade array-added below the if / end if
        Else
            If aryGrades(0, 0) = "" Then
                'zero position hasn't been used yet, use it first
                J = 0
            Else
                J = UBound(aryGrades, 2) + 1
            End If
            
            ReDim Preserve aryGrades(3, J)
            aryGrades(0, J) = CStr(gudtPDLine(I).GradeID)
            aryGrades(1, J) = GetGradeData(CStr(gudtPDLine(I).GradeID), "ID", "HHAID")
            aryGrades(2, J) = 0
            aryGrades(3, J) = 0
            'new grade added to the array
        End If
        
        'Now add the gross/net to the existing (if ffound=false just adding to zero)
        aryGrades(2, J) = Round(CDbl(aryGrades(2, J)) + gudtPDLine(I).PDLineGross, 2)
        aryGrades(3, J) = Round(CDbl(aryGrades(3, J)) + gudtPDLine(I).PDLineNet, 2)
            
        'Continue to the next board/check grade/add footages to totals
    Next
    
   'Now the Grade Breakdown / should be correct / now just need to sort by sort order in grade table
   'Return a List to update the PD Header Field (Like Chain Tally is Structured
    strReturnTallyDetail = ""
    
    
    'aryReturnGrades is not just an array it's really the same thing however Can't just be an array of strings,
        '''appforge limitation / can't pass arrays in modules so this is a custom type that just holds an array of strings
        
    ReDim aryReturnGrades.ary(3, UBound(aryGrades, 2))
    
    For I = 0 To UBound(garyGrade, 2)
        For J = 0 To UBound(aryGrades, 2)
            If SC(aryGrades(0, J), garyGrade(0, I)) = True Then 'this is the grade, move to same position in the return array
                aryReturnGrades.ary(0, J) = aryGrades(0, J)
                aryReturnGrades.ary(1, J) = aryGrades(1, J)
                aryReturnGrades.ary(2, J) = aryGrades(2, J)
                aryReturnGrades.ary(3, J) = aryGrades(3, J)
                
                'Build the String (same format as chain tally) to return to the Header Record
                'Tilda ~ separates grades and */* separates gradeid from the footage
                'Example GradeID=5 and Bft for that Grade=255 & GradeID=8 and Grade Bft=677
                '5*/*255~8*/*677
                Dim strGradeFootage As String
                strGradeFootage = "0"
                
                If fisGreen = True Then
                    strGradeFootage = CStr(aryReturnGrades.ary(2, J))
                Else
                    strGradeFootage = CStr(aryReturnGrades.ary(3, J))
                End If
                
                If SC(strReturnTallyDetail, "") = False Then strReturnTallyDetail = strReturnTallyDetail & "~"
                
                strReturnTallyDetail = strReturnTallyDetail & aryReturnGrades.ary(0, J) & "*/*" & strGradeFootage
                'END BUILD TALLY DETAIL STRING FOR HEADER BUNDLE ROW/FIELD
                
                'exit inner for loop / goto next grade
                Exit For
            Else
                'keep looking for grade/match in control array (garygrade) which is sorted in teh proper grade order
            End If
        Next
    Next
    'END of Build aryReturnGrades Array in the Proper Sort Order / Build Grade Footage String Values for ChainTally Header Fields
    'Now the aryReturnGrades array is recalculated, and the TAllyDetail string is built/ready to return to calling module
    
    Bundle_GetMultiGradeTotals_GlobalPDVariables = True
    
    Exit Function
ErrorHandler:
    
    MsgBox "Error in Bundle_GetMultiGradeTotals_GlobalPDVariables " & Err.Number & "-" & Err.Description
    Bundle_GetMultiGradeTotals_GlobalPDVariables = False
    Exit Function
End Function


Public Sub WriteEventLog(lngID As Long, strLogFileNameSuffix_TypeName As String, strLogSubType As String, lngUser As Long, dteDateTime As Date, strLogDesc As String)
    Dim cFile As CFileManager
    Dim cFileLog As CFileTextWritable
    Dim intError As Integer
    Dim strSDCardLogFileName As String
    Set cFile = New CFileManager
    Dim cFileSDCard As CFileTextWritable
    Dim strYearMonthDay As String
    Dim strLogFileName As String
    
    'for the file name
    strYearMonthDay = CStr(Year(Now)) & FixStringWidth(CStr(Month(Now)), 2, True, True) & FixStringWidth(CStr(day(Now)), 2, True, True)
On Error GoTo ErrorHandler

    intError = 10
    If strLogFileName = "" Then
        strLogFileName = gstrPDBPath & "\log_" & strLogFileNameSuffix_TypeName & "_" & strYearMonthDay & ".txt"
        strSDCardLogFileName = gstrStorageCardPath
        If Right(Trim(strSDCardLogFileName), 1) <> "\" Then strSDCardLogFileName = Trim(strSDCardLogFileName) & "\"
        strSDCardLogFileName = strSDCardLogFileName & "log_" & strLogFileNameSuffix_TypeName & "_" & strYearMonthDay & ".txt"
        
        intError = 15
    Else
        If Right(Trim(strSDCardLogFileName), 1) <> "\" Then Trim (strSDCardLogFileName = strSDCardLogFileName) & "\"
        
        strSDCardLogFileName = strSDCardLogFileName & "log_" & strLogFileNameSuffix_TypeName & "_" & strYearMonthDay & ".txt"
    End If
    intError = 17
''*************WRITE LOG TO PROGRAM FILES try to open first if exists/append to it ... or create if it doesn't exist
    

On Error GoTo TrySDCardWriteFile
        
TrySDCardWriteFile:
intError = 70
On Error GoTo TryToCreateIfNotOpenedSDCard
        Set cFileLog = cFile.OpenAsText(strSDCardLogFileName, afFileModeCreate)
        
intError = 80 'If this fails it will goto one below to try and create file instead of open it (if it doesn't already exist)
TryToCreateIfNotOpenedSDCard:
    'open file failed on sd card so create one

On Error GoTo ErrorHandler
    If intError = 70 Then ' the open file command didn't work, try to create it instead!
        Set cFileLog = cFile.OpenAsText(strSDCardLogFileName, afFileModeCreate)
    End If
    intError = 80
    Dim strFileDataCurrent As String
    
    strFileDataCurrent = ""
On Error GoTo SkipFileDataCurrentRead
    strFileDataCurrent = cFileLog.ReadToEnd
    
SkipFileDataCurrentRead:
On Error GoTo ErrorHandler
    strFileDataCurrent = strFileDataCurrent & vbCrLf & CStr(lngID) & "," & strLogFileNameSuffix_TypeName & "," & strLogSubType & "," & lngUser & "," & CStr((Now)) & "," & strLogDesc
    cFileLog.WriteLine strFileDataCurrent
    
    
    Set cFileLog = Nothing ' close the file
    Set cFile = Nothing
    
    Exit Sub
ErrorHandler:

'If it failed to write to the sd card then write to program files:
''Program Files Log File
On Error GoTo TryToCreateIfNotOpened
    intError = 20
    Set cFileLog = cFile.OpenAsText(strLogFileName, afFileModeOpen)
    intError = 30

TryToCreateIfNotOpened:
    'if the open above files then try to create it instead...it never hits error=30 above if fails.
On Error GoTo ErrorHandler2 ' if below / create file in program files fails then try the sd card file write .. also trys if works.
    If intError = 20 Then ' the open file command didn't work, try to create it instead!
        Set cFileLog = cFile.OpenAsText(strLogFileName, afFileModeCreate)
    End If
    intError = 40
    cFileLog.WriteLine CStr(lngID) & "," & strLogFileNameSuffix_TypeName & "," & strLogSubType & "," & lngUser & "," & CStr((Now)) & "," & strLogDesc
    Set cFileLog = Nothing ' close the file
    Exit Sub

'End Program Files Write Line
''***********************If above completes or failes try to also write to sd card so never deleted
'************************
'************************
    Exit Sub

On Error GoTo ErrorHandler2
    MsgBox "Error in WriteEventLog (Could not write Delete to SD Card (" & strSDCardLogFileName & ") or Program Files: (" & strLogFileName & ") - Type=" & strLogFileNameSuffix_TypeName & " SubType= " & strLogSubType & " Desc = " & strLogDesc & vbCrLf & "Error # (elIT#=" & intError & ") VB #= " & Err.Number & "-" & Err.Description
    Set cFile = Nothing
    Set cFileLog = Nothing
    

    Exit Sub
ErrorHandler2:
    
    Exit Sub
End Sub

Public Function fGetBEGroup(lngSpeciesID As Long, lngWidth As Long, lngGradeID As Long, lngThicknessID As Long, strGrade As String, lngLength As Long) As Integer
    Dim I As Integer
    Dim strLength As String
    Dim lngLow As Long
    Dim lnghigh As Long
    Dim lngBELengthID As Long
    Dim fFound As Boolean
    
Dim cbo As AFComboBox

'''    For i = 0 To UBound(garyBE)
'''        lngBELengthID = garyBE(i).LengthID
'''
'''        If garyBE(i).Active = 1 Then
'''            If garyBE(i).SpeciesID = lngSpeciesID And garyBE(i).GradeID = lngGradeID And garyBE(i).ThicknessID = lngThicknessID And lngGradeID = garyBE(i).GradeID Then
'''                'Check the length and width as necessary
'''                Call AppForge_Split(garyBE(i).width, pstrSplit, "-")
'''                If UBound(pstrSplit.ary) = 1 Then
'''                    If lngWidth >= CLng(pstrSplit.ary(0)) And lngWidth <= CLng(pstrSplit.ary(1)) Then
'''                        'Still in the running, now check the length.
'''                        strLength = GetLengthData(CStr(lngBELengthID), "ID")
'''
'''                        Call AppForge_Split(strLength, pstrSplit, ",")
'''                        If UBound(pstrSplit.ary) = 5 Then
'''                            If lngLength >= CLng(pstrSplit.ary(4)) And lngLength <= CLng(pstrSplit.ary(5)) Then
'''                                'Jackpot....this is the right one
'''                                ''This is it...Return this Index (I) to use with GaryBE(I) in the calling function
'''                                fFound = True
'''                                Exit For
'''                            End If
'''                        End If
'''
'''
'''                    End If
'''                End If
'''            End If
'''
'''        End If
'''    Next
    
    If fFound = True Then
        fGetBEGroup = I
    Else
        fGetBEGroup = -1
    End If
    
End Function

Public Function GetWeekTotalFootage(dteNow As Date) As Double
    
    Dim fComplete As Boolean
    Dim dblFootage As Double
    Dim dteStart As Date, dteEnd As Date
    Dim intDiff As Integer
    Dim udt As tPDRecord
    
'''    If Weekday(dteNow) > 1 Then
'''        intDiff = Weekday(dteNow) - 1
'''        dteStart = CDate(FormatDateTime(DateAdd("d", -1 * intDiff, dteNow), vbShortDate))
'''    Else
'''        dteStart = CDate(FormatDateTime(dteNow, vbShortDate))
'''    End If
'''
'''    dteEnd = DateAdd("d", 7, dteStart)
'''
'''
'''
'''    db = PDBOpen(Byfilename, gstrPDBPath & "\pd", 0, 0, 0, 0, afModeReadWrite)
'''
'''    If db = 0 Then
'''        MsgBox "Unable to open production database"
'''        Exit Function
'''    End If
'''
'''    PDBSetSortFields db, 10
'''    PDBMoveLast db
'''
'''    fComplete = False
'''    Do Until fComplete = True
'''        PDBReadRecord db, VarPtr(udt)
'''        Debug.Print udt.PDRoughGradingDate
'''        If udt.GreenDry = 1 Then
'''            If dteStart > udt.PDRoughGradingDate Then
'''                fComplete = True
'''            ElseIf dteStart <= udt.PDRoughGradingDate And dteEnd >= udt.PDRoughGradingDate Then
'''                'Count this one
'''                dblFootage = dblFootage + udt.PDTotalNetBFM
'''            End If
'''        ElseIf dteStart > udt.PDRoughGradingDate Then
'''            fComplete = True
'''        End If
'''        PDBMovePrev db
'''    Loop
'''
'''    PDBClose db
'''    db = 0
'''    GetWeekTotalFootage = dblFootage
'''    gfWeekTotal = True
    
End Function

Public Sub UpdateTallyKeyList(intKeypad As Integer, Optional start As Integer = 0, Optional spe As String = "", Optional speid As String = "")
    Dim intRows As Integer
    Dim I As Integer, J As Integer
    
    
    If start = 0 And UBound(gTallyKey, 2) > 1 Then Exit Sub 'already setup
    
    If intKeypad = 36 Then
        intRows = 15
    End If
    Dim jk As Integer
    jk = 0
    If start <> 0 Then jk = 1
  ReDim Preserve gTallyKey(TallyKey.Last, UBound(gTallyKey, 2) + intRows + jk)
  
    For I = 0 + start To intRows + start
        For J = 0 To TallyKey.Last
            gTallyKey(J, I) = ""
        Next
        gTallyKey(TallyKey.GradeID, I) = 0
        gTallyKey(TallyKey.ThicknessID, I) = 0
        gTallyKey(TallyKey.Species, I) = spe
        gTallyKey(TallyKey.SpeciesID, I) = speid
    Next
    
    
    'Zero Key
    gTallyKey(TallyKey.KeyCode, 0 + start) = "48"
    gTallyKey(TallyKey.KeyText, 0 + start) = "0"
    
    '1-9 Keys
    J = 1 + start
    For I = 49 To 57
        gTallyKey(TallyKey.KeyCode, J) = CStr(I)
        gTallyKey(TallyKey.KeyText, J) = CStr(J - start)
        J = J + 1
    Next
    
    'Function Keys F1-F5
    For I = 112 To 116
        gTallyKey(TallyKey.KeyCode, J) = CStr(I)
        gTallyKey(TallyKey.KeyText, J) = "F" & CStr(J - 9 - start)
        J = J + 1
    Next
    
    'Function Keys F6 - it is out of sequence on HH
    gTallyKey(TallyKey.KeyCode, J) = 96
    gTallyKey(TallyKey.KeyText, J) = "F6"
        
'Other Keys that could be used on 32 Key Unit
'''    J = J + 1
'''    'SP Key on 32 Key HH
'''    gTallyKey(TallyKey.KeyCode, J) = 32
'''    gTallyKey(TallyKey.KeyCode, J) = "SP"
'''
'''    J = J + 1
'''    'BKSP Key on 32 Key HH
'''    gTallyKey(TallyKey.KeyCode, J) = 8
'''    gTallyKey(TallyKey.KeyCode, J) = "BKSP"
'''
'''    J = J + 1
'''    'CTRL Key on 32 Key HH
'''    gTallyKey(TallyKey.KeyCode, J) = 17
'''    gTallyKey(TallyKey.KeyCode, J) = "CTRL"
'''
    
    
    Exit Sub
    
End Sub
Public Sub LoadTallyKey()
    
    Dim I As Integer, J As Integer
    Dim aryTallyKey() As tTallyKeyRecord

On Error GoTo ErrorHandler

    OpenTallyKeyDatabase     ' ltas 8.29.2006
    Call ErrorCheck(dbTallyKey)
    If dbTallyKey = 0 Then
        MsgBox "Unable to open Grade Setup database"
        Exit Sub
    End If
    PDBMoveFirst dbTallyKey
    I = -1
    
    ReDim aryTallyKey(0)
    aryTallyKey(0).Last = "-1"
    
    Do Until PDBEOF(dbTallyKey)
        I = I + 1
        If I = 0 Then
            'do nothing
        Else
            ReDim Preserve aryTallyKey(I)
        End If
        
        Call ReadTallyKeyRecord(aryTallyKey(I))
        PDBMoveNext dbTallyKey
    Loop
            
    PDBClose dbTallyKey
    dbTallyKey = 0
    If UBound(aryTallyKey) < 1 Then Exit Sub
    ReDim gTallyKey(TallyKey.Last, UBound(aryTallyKey))
    For I = 0 To UBound(aryTallyKey)
      '  For j = 0 To UBound(gTallyKey)
          '  If gTallyKey(TallyKey.KeyCode, j) = aryTallyKey(i).KeyCode Then
                gTallyKey(TallyKey.Species, I) = aryTallyKey(I).Species
                gTallyKey(TallyKey.SpeciesID, I) = aryTallyKey(I).SpeciesID
                gTallyKey(TallyKey.Grade, I) = aryTallyKey(I).Grade
                gTallyKey(TallyKey.GradeHHID, I) = aryTallyKey(I).GradeHHID
                gTallyKey(TallyKey.GradeID, I) = aryTallyKey(I).GradeID
                gTallyKey(TallyKey.KeyCode, I) = aryTallyKey(I).KeyCode
                gTallyKey(TallyKey.KeyText, I) = aryTallyKey(I).KeyText
                gTallyKey(TallyKey.Last, I) = aryTallyKey(I).Last
                gTallyKey(TallyKey.LoadAID, I) = aryTallyKey(I).LoadAID
                gTallyKey(TallyKey.thickness, I) = aryTallyKey(I).thickness
                gTallyKey(TallyKey.ThicknessHHID, I) = aryTallyKey(I).ThicknessHHID
                gTallyKey(TallyKey.ThicknessID, I) = aryTallyKey(I).ThicknessID
           '     Exit For
         '   End If
       ' Next
    Next
    
    'Now check the validity of the Grades Assigned to Keys in the TallyKey.pdb to help troubleshoot errors (only when gradekeys=1)
    If gintGradeKeys = 0 Then
        Call Utility_Search("USEGRADEKEYS", "", "NAME", True)
        gintGradeKeys = CInt(gudtUtility(0).UtilityValueLong)
    End If
    
    If gintGradeKeys = 1 Then
        If gfTallyKeyWarningAlreadyShown = True Then
            'only show it once per launch/running so not constantly annoying people not doing chain tallies!
            Exit Sub
        
        Else
            'check grade ids in tallykey.pdb and make sure they are valid for this handhelds grade.pdb table
            
            Dim strMsg As String
            Dim strTempGradeHHAID As String
            strMsg = ""
            
            For I = 0 To UBound(gTallyKey, 2)
                If CLng(gTallyKey(TallyKey.GradeID, I)) > 0 Then
                    strTempGradeHHAID = ""
                    strTempGradeHHAID = tcu(GetGradeData(CStr(gTallyKey(TallyKey.GradeID, I)), "ID", "HHAID"))
                    
                    If InStr(strTempGradeHHAID, "INVALID") > 0 Then
                        'There is a gradeid in the tally key database that isn't valid
                        If SC(strMsg, "") = True Then
                            strMsg = "The Following Grade IDs/Keys are Invalid in the Chain Tally setup:"
                        End If
                        strMsg = strMsg & vbCrLf & "GradeID " & gTallyKey(TallyKey.GradeID, I) & " in TallyKey HH Database is Assigned to Key=" & gTallyKey(TallyKey.KeyText, I) & " And Is Invalid!"
                    End If
                End If
            Next
            
            If SC(strMsg, "") = False Then
                MsgBox strMsg
            End If
            gfTallyKeyWarningAlreadyShown = True ' update the flag so it's not continuously checked/warned-once per opening mLIMBS
        End If
    End If
    
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in LoadTallyKey_basCommon " & vbCrLf & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub

Public Sub SaveTallyKey()
    
    Dim I As Integer, J As Integer
    Dim aryTallyKey() As tTallyKeyRecord
    Dim udt As tTallyKeyRecord
    Dim fInsert As Boolean
    
    Dim FSO As CFileManager
     Set FSO = New CFileManager
     FSO.DeleteFile (gstrPDBPath & "\TallyKey.pdb")

        CreateDatabasePDB dbPD, "TallyKey", TallyKey_Schema
        PDBClose dbTallyKey
    
    
    OpenTallyKeyDatabase     ' ltas 8.29.2006
    Call ErrorCheck(dbTallyKey)
    If dbTallyKey = 0 Then
        MsgBox "Unable to open Grade Setup database"
        Exit Sub
    End If
    PDBMoveFirst dbTallyKey
        
    ReDim aryTallyKey(UBound(gTallyKey, 2))
    
    For I = 0 To UBound(gTallyKey, 2)
        aryTallyKey(I).Species = gTallyKey(TallyKey.Species, I)
        aryTallyKey(I).SpeciesID = gTallyKey(TallyKey.SpeciesID, I)
        aryTallyKey(I).Grade = gTallyKey(TallyKey.Grade, I)
        aryTallyKey(I).GradeHHID = gTallyKey(TallyKey.GradeHHID, I)
        aryTallyKey(I).GradeID = gTallyKey(TallyKey.GradeID, I)
        aryTallyKey(I).KeyCode = gTallyKey(TallyKey.KeyCode, I)
        aryTallyKey(I).KeyText = gTallyKey(TallyKey.KeyText, I)
        aryTallyKey(I).Last = gTallyKey(TallyKey.Last, I)
        aryTallyKey(I).LoadAID = gTallyKey(TallyKey.LoadAID, I)
        aryTallyKey(I).thickness = gTallyKey(TallyKey.thickness, I)
        aryTallyKey(I).ThicknessHHID = gTallyKey(TallyKey.ThicknessHHID, I)
        aryTallyKey(I).ThicknessID = gTallyKey(TallyKey.ThicknessID, I)
    Next
    
    For I = 0 To UBound(aryTallyKey)
       ' PDBFindRecordByField dbTallyKey, 1, aryTallyKey(i).KeyCode
       ' If PDBGetLastError(dbTallyKey) = -3 Or PDBGetLastError(dbTallyKey) = -7 Then
            PDBCreateRecordBySchema dbTallyKey
       ' Else
       '     ReadTallyKeyRecord udt
       '     PDBEditRecord dbTallyKey
       ' End If
        udt.Species = aryTallyKey(I).Species
        udt.SpeciesID = aryTallyKey(I).SpeciesID
        udt.Grade = aryTallyKey(I).Grade
        udt.GradeHHID = aryTallyKey(I).GradeHHID
        udt.GradeID = aryTallyKey(I).GradeID
        udt.KeyCode = aryTallyKey(I).KeyCode
        udt.KeyText = aryTallyKey(I).KeyText
        udt.Last = aryTallyKey(I).Last
        udt.LoadAID = aryTallyKey(I).LoadAID
        udt.thickness = aryTallyKey(I).thickness
        udt.ThicknessHHID = aryTallyKey(I).ThicknessHHID
        udt.ThicknessID = aryTallyKey(I).ThicknessID
        fInsert = WriteTallyKeyRecord(udt)
        PDBUpdateRecord dbTallyKey
    Next
            
    PDBClose dbTallyKey
    dbTallyKey = 0
    
End Sub

Public Sub BackupTimer(lblBackup As AFLabel, lblAutoBackup As AFLabel)
    intTimer = intTimer + 1
        
On Error GoTo ErrorHandler

    If gfOther = True Then Exit Sub
    
    lblAutoBackup.Visible = False
    lblAutoBackup.Refresh
    
    If intTimer < 2 Then
        lblAutoBackup.Visible = False
        lblAutoBackup.Refresh
    ElseIf intTimer = 2 Then
        lblAutoBackup.Caption = "AutoBackup: 2min"
        lblAutoBackup.Visible = True
        lblAutoBackup.Refresh
    ElseIf intTimer = 2 Then
        lblAutoBackup.Caption = "AutoBackup: 1.5min"
        lblAutoBackup.Visible = True
        lblAutoBackup.Refresh
    ElseIf intTimer = 3 Then
        lblAutoBackup.Caption = "AutoBackup: 1min"
        lblAutoBackup.Visible = True
        lblAutoBackup.Refresh
    ElseIf intTimer = 4 Then
        lblAutoBackup.Caption = "AutoBackup: 30sec"
        lblAutoBackup.Visible = True
        lblAutoBackup.Refresh
    ElseIf intTimer = 5 Then
        lblAutoBackup.Caption = "INITIATING..."
        lblAutoBackup.Visible = True
        lblBackup.ZOrder 0
        lblBackup.Visible = True
        lblBackup.Refresh
        lblAutoBackup.Refresh
        intTimer = 0
        If gfBackupNeeded = True Then frmBackupShow.Show
        lblBackup.Visible = False
        lblAutoBackup.Visible = False
        lblBackup.Refresh
        lblAutoBackup.Refresh
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in BackupTimer: " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub
Public Function BackuptoSD(lbl As AFLabel, Optional fManualBackup As Boolean, Optional strSDFolder As String, _
    Optional fNoIPSM As Boolean = False, Optional strSkipOption As String, Optional fUploadBackup As Boolean) As Boolean
        
    Dim FSO As CFileManager
    Dim fSDBackup As Boolean
    Dim strQtrHour As String
    Dim strStorageCardPath As String
    Dim strMonth As String
    Dim strDay As String
    Dim strHour As String
    Dim I As Integer
    Dim intError As Integer
    Dim strBackupType As String
    Dim lngBackupID As Long
    
On Error GoTo ErrorHandler
    If gfSMTallyFormRunning = True Then 'NO BACKUPS FROM SM TALLY FORM CURRENTLY IF IT"S OPEN AT ALL
        Exit Function
    End If
    
    Set FSO = New CFileManager
    
intError = 1

    gintBackupCountCT = gintBackupCountCT + 1
    If gintBackupCountCT > 30 Then gintBackupCountCT = 1
                
    'MsgBox "Backup about to begin..."
    
    If Minute(Now) > 0 And Minute(Now) < 16 Then strQtrHour = "01"
    If Minute(Now) > 15 And Minute(Now) < 31 Then strQtrHour = "02"
    If Minute(Now) > 30 And Minute(Now) < 46 Then strQtrHour = "03"
    If Minute(Now) > 45 And Minute(Now) < 61 Then strQtrHour = "04"
    
intError = 10

    strMonth = CStr(Month(Now))
    strDay = CStr(day(Now))
    
    If Len(strMonth) = 1 Then
        strMonth = "0" & strMonth
    End If
    
    If Len(strDay) = 1 Then
        strDay = "0" & strDay
    End If
    
intError = 15
    
    Call Utility_Search("BACKUPTYPE", "", "NAME", True, True, False, "")
    strBackupType = Trim(UCase(gudtUtility(0).UtilityValueText))
    
    Call Utility_Search("BACKUPCOUNT", "", "NAME", True, True, False, "")
    If IsNumeric(gudtUtility(0).UtilityValueText) = False Then
        lngBackupID = 1
    Else
        lngBackupID = CLng(gudtUtility(0).UtilityValueText)
        If lngBackupID > 6 Then lngBackupID = 1
    End If
    
    If strBackupType = "DISABLED" Then
        lbl.Caption = "BACKUPS DISABLED"
        Unload frmBackupShow
        Exit Function
    End If
    
'    If fManualBackup = True Or gintBackupCountCT Mod 2 = 0 Then
        fSDBackup = True
        If strBackupType <> "LIMITED6" Then
            If fUploadBackup = True Then
                strSDFolder = "\Upload_Minute_" & CStr(Minute(Now)) & "_Day_"
            Else
                strSDFolder = "\"
            End If
            
            strSDFolder = strSDFolder & strMonth & "_" & strDay & "_" & strHour & "_" & strQtrHour & "_v" & CStr(App.Major) & CStr(App.Minor) & CStr(App.Revision)
        Else
            strSDFolder = "\mLIMBS_" & CStr(lngBackupID)
        End If
        
        
'    End If
    
    If fManualBackup = True Then fSDBackup = True
    
    lbl.BackColor = vbRed
    lbl.ZOrder 0
    lbl.Caption = "Auto - Backup ... Please Wait - 15"
    lbl.Refresh
    lbl.Visible = True
    lbl.Refresh
    

intError = 25

    If strSDFolder <> "" Then
        
        strStorageCardPath = gstrStorageCardPath & strSDFolder
        Call func_CreateDirectory(strStorageCardPath, FSO)
        Call FSO.OpenDirectory(strStorageCardPath, True)
    Else
        strStorageCardPath = gstrStorageCardPath
    End If
    
    #If AppForge Then 'continue
    #Else
        lbl.Visible = False
        lbl.Refresh
        Unload frmBackupShow
        
        Set FSO = Nothing
        Exit Function
   #End If
    
    
    
    closealldatabases
    LoadPDBArray
intError = 30

    For I = 0 To UBound(garyPDB)
        
        intError = intError + 1
        If Trim(UCase(garyPDB(I))) <> "" Then
            
            If (strSkipOption = "RHRL" And tcu(garyPDB(I)) = "RH.PDB") Or (strSkipOption = "RHRL" And garyPDB(I) = "RL.PDB") Or Trim(UCase(garyPDB(I))) = Trim(UCase("PDINV.pdb")) Then
             'PDINV   On Error GoTo SkipFile
             'PDINV   If fSDBackup = True Then fso.CopyFile gstrIPSM & "\" & garyPDB(I), strStorageCardPath & "\" & garyPDB(I), True
                '
            Else
                On Error GoTo ErrorHandler
                
                If fSDBackup = True Then FSO.CopyFile gstrPDBPath & "\" & garyPDB(I), strStorageCardPath & "\" & garyPDB(I), True
                If fNoIPSM = False Then FSO.CopyFile gstrPDBPath & "\" & garyPDB(I), gstrIPSM & "\" & garyPDB(I), True
                lbl.Caption = "Auto - Backup ... Please Wait - " & I & " of " & CStr(garyPDB(I)) & " To: " & strSDFolder
                lbl.Refresh
            End If
        End If
SkipFile:
    Next
    lbl.Refresh
    
intError = 100

    lbl.Caption = "Auto - Backup Complete"
    lbl.Refresh
    
    lbl.Visible = False
    lbl.Refresh
    Set FSO = Nothing
    
    gfBackupNeeded = False
    
    BackuptoSD = True
    
    If strBackupType = "LIMITED6" Then
        Call Utility_Search("BACKUPCOUNT", CStr(lngBackupID + 1), "NAME", True, False, True, "")
    End If
    
    Call Utility_Search("BACKUPLASTDATE", CStr(Now), "NAME", True, True, True, "")
    
intError = 150
    Unload frmBackupShow
    
    Exit Function

ErrorHandler:

On Error GoTo ErrorHandler2
    BackuptoSD = False
    If frmBackupShow.Visible = True Then Unload frmBackupShow
    If I <= UBound(garyPDB) Then
        MsgBox intError & "-Error Backing Up:" & garyPDB(I) & " to " & gstrStorageCardPath & strSDFolder
    Else
        MsgBox intError & " -  Error During Backup (" & gstrStorageCardPath & strSDFolder & ") - " & Err.Number & Err.Description & " - BACKUP NOT COMPLETED.  Please Exit System and Restart to Try again!"
    End If
    
    lbl.Visible = False
    lbl.Refresh
    Exit Function

ErrorHandler2:
    MsgBox Err.Number & " " & Err.Description & " " & "- Error During Backup "
    Exit Function
    
End Function
Public Sub DeleteUtilityRecord(grd As AFGrid)
    
On Error GoTo ErrorHandler

    If grd.Rows <> 0 Then
        OpenUtilityDatabase
        If DeleteUtilityRecordByName(grd.TextMatrix(grd.Row, 1)) = True Then
            grd.RemoveItem (grd.Row)
        Else
            MsgBox "Could not delete Utility Record With SettingName=" & grd.TextMatrix(grd.Row, 1)
        End If
    End If
    
    Exit Sub
ErrorHandler:

    MsgBox "Utility Record/Setting Failed to delete!!!" & vbCrLf & "Error in DeleteUtilityRecord(grd as AFGrid) " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub
Public Function DeleteUtilityRecordByName(strUtilityName As String) As Boolean
    Dim fDeleteComplete As Boolean
On Error GoTo ErrorHandler
    fDeleteComplete = False
    
    OpenUtilityDatabase
    Call PDBFindRecordByField(dbUtility, tUtilityDatabaseFields.UtilityName_Field, strUtilityName)
        
        If PDBGetLastError(dbUtility) = 0 Then
            PDBDeleteRecordEx dbUtility, afDeleteModeRemove
            fDeleteComplete = True
        End If
        CloseUtilityDatabase
    
    
    DeleteUtilityRecordByName = fDeleteComplete
    
    Exit Function
ErrorHandler:
    DeleteUtilityRecordByName = False
    MsgBox "Utility Record/Setting Failed to delete!!!" & vbCrLf & "Error in DeleteUtilityRecordbyName(Name=" & strUtilityName & ") " & Err.Number & "-" & Err.Description
    Exit Function
    
End Function

Public Sub RestoreBackup2009(Optional fld As String = "", Optional lbl As AFLabel, Optional fManualRestore As Boolean)
    Dim FSO As CFileManager
    Dim intError As Integer
    Dim strPath As String
    Dim I As Integer, J As Integer
    Dim fDir As CDirectory
    
On Error GoTo ErrorHandler
    
    Set FSO = New CFileManager
    
    If fManualRestore = False Then
        If MsgBox("System Restore Needed?  More than 5 PDB Files are Missing would you like to restore now?", vbYesNo) = vbNo Then
            End
        End If
    End If
    
    
intError = 1
    If fld = "" Then
        If gSymbol = True Then
            strPath = "\Platform"
        Else
            strPath = "\IPSM"
            If gf9900 = True Then
                'strPath = "\IPSM\mLIMBSBackup"
                strPath = "\IPSM"
            End If
        End If
    Else
        If gSymbol = True Then
            strPath = "\Platform"
        Else
            strPath = "\storage card"
        End If
        strPath = strPath & "\" & fld
    End If
    
intError = 2
    Call LoadPDBArray
intError = 3
    
    Set fDir = FSO.OpenDirectory(strPath, True)
    Set fDir = Nothing
    '
    For I = 0 To UBound(garyPDB)
        intError = intError + 1
        If Trim(garyPDB(I)) <> "" And garyPDB(I) <> "PDINV" Then
            lbl.Caption = "Restoring: " & Replace(garyPDB(I), ".pdb", "")
            lbl.Refresh
            
            FSO.CopyFile strPath & "\" & garyPDB(I), gstrPDBPath & "\" & garyPDB(I), True
            
            lbl.Caption = "Restoring: " & Replace(garyPDB(I), ".pdb", "") & " Successful"
            lbl.Refresh
        End If
    Next
    
    Dim aryFiles(6) As String
        
    aryFiles(0) = "CHAINTAG.lbl"
    aryFiles(1) = "ENDTALLY.lbl"
    aryFiles(2) = "ReceiveTag.lbl"
    aryFiles(3) = "ReceiveTagDetail.lbl"
    aryFiles(4) = "ReceiveTicket.lbl"
    aryFiles(5) = "BTTag1.lbl"
    aryFiles(6) = "BTTag2.lbl"
    
    For I = 0 To UBound(aryFiles)
intError = 950
        FSO.CopyFile strPath & "\" & aryFiles(I), gstrPDBPath & "\" & aryFiles(I), True
GotoNextFile:
    Next
    
    Exit Sub

ErrorHandler:
    If intError = 950 Then
        GoTo GotoNextFile
    Else
        MsgBox intError & " - Source: " & strPath & "\" & garyPDB(I) & " Failed to Copy to File to: " & gstrPDBPath & "\" & garyPDB(I) & "  Soft Reset and Try again, Contact Tech Support if Continues at 740.401.0720"
    End If
    End
    
End Sub

Public Function BackupgaryCTLinetoSD(lbl As AFLabel, Optional fPDOnly As Boolean) As Boolean
    Dim FSO As CFileManager
    Dim strFolder As String
    Dim fSDBackup As Boolean
On Error GoTo ErrorHandler
    Set FSO = New CFileManager
    
        
    gintBackupCTLine = gintBackupCTLine + 1
    
    If gintBackupCTLine = 25 Then gintBackupCTLine = 0
        
    Call Utility_Search("BACKUPCOUNT-CT", CStr(gintBackupCountCT), "NAME", True, True)
    
    
    If gSymbol = True Then
        strFolder = "\Platform\CTLine_" & CStr(gintBackupCTLine)
    Else
        strFolder = "\Storage Card\CTLine_" & CStr(gintBackupCTLine)
    End If
    
    If fPDOnly = True Then
        strFolder = Replace(strFolder, "CTLine_", "PD_PDLine_")
    End If
    
    
    Call closealldatabases
        
    lbl.BackColor = vbGreen
    lbl.ZOrder 0
    lbl.Caption = "Load Auto - Backup ... Please Wait - 3"
    lbl.Refresh
    lbl.Visible = True
    lbl.Refresh
    
    #If AppForge Then
        FSO.OpenDirectory strFolder, True
    #Else
        BackupgaryCTLinetoSD = True
        lbl.Caption = "SKIPPING - DEVELOPING/RUNNING"
        Exit Function
    #End If
    
    lbl.Caption = "Load Auto - Backup ... Please Wait - 2"
    lbl.Refresh
    lbl.Visible = True
    lbl.Refresh
    
    If fPDOnly = False Then
        FSO.CopyFile gstrPDBPath & "\CTLine.pdb", strFolder & "\CTLine.pdb", True
        lbl.Caption = "Load Auto - Backup ... Please Wait - 1"
        lbl.Refresh
        
        FSO.CopyFile gstrPDBPath & "\Load.pdb", strFolder & "\Load.pdb", True
        lbl.Caption = "Load Auto - Backup ... Please Wait - 1"
        lbl.Refresh
    End If
    
    FSO.CopyFile gstrPDBPath & "\pd.pdb", strFolder & "\pd.pdb", True
    lbl.Caption = "Load Auto - Backup ... Please Wait - 1"
    lbl.Refresh
    
    FSO.CopyFile gstrPDBPath & "\pdline.pdb", strFolder & "\pdline.pdb", True
    lbl.Caption = "Load Auto - Backup ... Please Wait - 1"
    lbl.Refresh
    
    lbl.Caption = "Load Auto - Backup Complete"
    lbl.Refresh
    
    lbl.Visible = False
    lbl.Refresh
    Set FSO = Nothing
    
    lbl.BackColor = vbRed
    BackupgaryCTLinetoSD = True
    Exit Function

ErrorHandler:
    MsgBox "Error During Load Backup - Clear SD Card and Try Again" & Err.Number & Err.Description & " - BACKUP NOT COMPLETED.  Please Exit System and Restart to Try again!"
    MsgBox "IF YOU IGNORE THIS ERROR YOU WILL POTENTIALLY LOSE LOAD INFORMATION!  CONTACT eLIMBS SUPPORT @ 740.401.0720"
    
    lbl.Visible = False
    lbl.Refresh
    BackupgaryCTLinetoSD = False
    Exit Function
    
End Function
Public Sub GetShipperName()
    
    Call Utility_Search("COMPANYNAME", "", "NAME", True)
    
    If gfDEMO = True Then
        gStrShipper = "DEMO - INVALID LICENSE"
    Else
        gStrShipper = gudtUtility(0).UtilityValueText
    End If
    
End Sub
    
Public Sub GetBTCommmPort()
    Call Utility_Search("BTCOMM", "", "NAME", True)
    If IsNumeric(gudtUtility(0).UtilityValueText) = True Then
        
        gintPrinterPort = CInt(gudtUtility(0).UtilityValueText)
    Else
        gintPrinterPort = 7
    End If
    
    
End Sub

Public Sub GetLocationDefault()
    Call Utility_Search("LOCATIONDEFAULT", "", "NAME", True, False)
    
    If gfDEMO = True Then
        gstrLastLocation = "DEMO - INVALID LICENSE"
    Else
        gstrLastLocation = gudtUtility(0).UtilityValueText
    End If

End Sub

Public Sub GetUpdateGradeKeyPref(strGetUpdate As String, intGradeKeys As Integer)
    Dim fInsert As Boolean
        
    Call Utility_Search("USEGRADEKEYS", "", "NAME", True)
    
    If gudtUtility(0).UtilityName = "" Then
        strGetUpdate = "NEW"
    End If
        
    If strGetUpdate = "GET" Then
        gintGradeKeys = CInt(gudtUtility(0).UtilityValueLong)
        Exit Sub
    ElseIf strGetUpdate = "UPDATE" Or strGetUpdate = "NEW" Then
    
        OpenUtilityDatabase     ' ltas 8.29.2006
        Call ErrorCheck(dbUtility)
        
        If dbUtility = 0 Then
            MsgBox "Unable to open Utility Setup database"
            Exit Sub
        End If
        
        PDBSetSortFields dbUtility, 0
        PDBMoveFirst dbUtility
        
        If strGetUpdate = "UPDATE" Then
            PDBFindRecordByField dbUtility, 0, gudtUtility(0).UtilityID
            'PDBFindRecordByField dbUtility, 1, gudtUtility(0).UtilityName
            
            PDBEditRecord dbUtility
            
            gudtUtility(0).UtilityValueText = CStr(intGradeKeys)
            gudtUtility(0).UtilityValueLong = CLng(intGradeKeys)
            
        Else 'NEW ' First Time ADD - Default to ON
            gudtUtility(0).UtilityName = "USEGRADEKEYS"
            gudtUtility(0).UtilityValueLong = 1
            gudtUtility(0).UtilityValueText = "1"
            gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
            PDBCreateRecordBySchema (dbUtility)
        End If
        
        fInsert = WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
                
        PDBClose dbUtility
        dbUtility = 0
        gintGradeKeys = CInt(gudtUtility(0).UtilityValueLong)

    Else
        Exit Sub
    End If
    
End Sub



Public Sub GetLoadAutoAdvancePref(strGetUpdate As String, intLoad As Integer)
    Dim fInsert As Boolean
        
    Call Utility_Search("LOADAUTOADVANCE", "", "NAME", True)
    
    If gudtUtility(0).UtilityName = "" Then
        strGetUpdate = "NEW"
    End If
        
    If strGetUpdate = "GET" Then
        gintLoadAutoAdvance = CInt(gudtUtility(0).UtilityValueLong)
        Exit Sub
    ElseIf strGetUpdate = "UPDATE" Or strGetUpdate = "NEW" Then
        OpenUtilityDatabase
        Call ErrorCheck(dbUtility)
        
        If dbUtility = 0 Then
            MsgBox "Unable to open Utility Setup database"
            Exit Sub
        End If
        
        PDBSetSortFields dbUtility, 0
        PDBMoveFirst dbUtility
        
        If strGetUpdate = "UPDATE" Then
            PDBFindRecordByField dbUtility, 0, gudtUtility(0).UtilityID
            'PDBFindRecordByField dbUtility, 1, gudtUtility(0).UtilityName
            
            PDBEditRecord dbUtility
            
            gudtUtility(0).UtilityValueText = CStr(intLoad)
            gudtUtility(0).UtilityValueLong = CLng(intLoad)
            
        Else 'NEW ' First Time ADD - Default to ON
            gudtUtility(0).UtilityName = "LOADAUTOADVANCE"
            gudtUtility(0).UtilityValueLong = 1
            gudtUtility(0).UtilityValueText = "1"
            gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
            PDBCreateRecordBySchema (dbUtility)
        End If
        
        fInsert = WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
                
        PDBClose dbUtility
        dbUtility = 0
        gintLoadAutoAdvance = CInt(gudtUtility(0).UtilityValueLong)

    Else
        Exit Sub
    End If
    
End Sub
Public Sub GetGradePercentagePref(strGetUpdate As String, intLoad As Integer)
    Dim fInsert As Boolean
        
    Call Utility_Search("GRADEPERCENTAGE", "", "NAME", True)
    
    If gudtUtility(0).UtilityName = "" Then
        strGetUpdate = "NEW"
    End If
        
    If strGetUpdate = "GET" Then
        gintGradePercentage = CInt(gudtUtility(0).UtilityValueLong)
        Exit Sub
    ElseIf strGetUpdate = "UPDATE" Or strGetUpdate = "NEW" Then
        OpenUtilityDatabase
        Call ErrorCheck(dbUtility)
        
        If dbUtility = 0 Then
            MsgBox "Unable to open Utility Setup database"
            Exit Sub
        End If
        
        PDBSetSortFields dbUtility, 0
        PDBMoveFirst dbUtility
        
        If strGetUpdate = "UPDATE" Then
            PDBFindRecordByField dbUtility, 0, gudtUtility(0).UtilityID
            PDBEditRecord dbUtility
            
            gudtUtility(0).UtilityValueText = CStr(intLoad)
            gudtUtility(0).UtilityValueLong = CLng(intLoad)
            
        Else 'NEW ' First Time ADD - Default to ON
            gudtUtility(0).UtilityName = "GRADEPERCENTAGE"
            gudtUtility(0).UtilityValueLong = 1
            gudtUtility(0).UtilityValueText = "1"
            gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
            PDBCreateRecordBySchema (dbUtility)
        End If
        
        fInsert = WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
                
        PDBClose dbUtility
        dbUtility = 0
        gintGradePercentage = CInt(gudtUtility(0).UtilityValueLong)

    Else
        Exit Sub
    End If
    
End Sub


Public Sub GetUpdateGradeKeyOrder(strGetUpdate As String, intSMFirst As Integer)
    Dim fInsert As Boolean
        
    Call Utility_Search("SMFIRST", "", "NAME", True)
    
    If gudtUtility(0).UtilityName = "" Then
        strGetUpdate = "NEW"
    End If
        
    If strGetUpdate = "GET" Then
        gintSMFirst = CInt(gudtUtility(0).UtilityValueLong)
        Exit Sub
    ElseIf strGetUpdate = "UPDATE" Or strGetUpdate = "NEW" Then
    
        OpenUtilityDatabase     ' ltas 8.29.2006
        Call ErrorCheck(dbUtility)
        
        If dbUtility = 0 Then
            MsgBox "Unable to open Utility Setup database"
            Exit Sub
        End If
        
        PDBSetSortFields dbUtility, 0
        PDBMoveFirst dbUtility
        
        If strGetUpdate = "UPDATE" Then
            PDBFindRecordByField dbUtility, 0, gudtUtility(0).UtilityID
            PDBEditRecord dbUtility
            
            gudtUtility(0).UtilityValueText = CStr(intSMFirst)
            gudtUtility(0).UtilityValueLong = CLng(intSMFirst)
            
        Else 'NEW ' First Time ADD - Default to ON
            gudtUtility(0).UtilityName = "SMFIRST"
            gudtUtility(0).UtilityValueLong = 0
            gudtUtility(0).UtilityValueText = "0"
            gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
            PDBCreateRecordBySchema (dbUtility)
        End If
        
        fInsert = WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
                
        PDBClose dbUtility
        dbUtility = 0
        gintSMFirst = CInt(gudtUtility(0).UtilityValueLong)

    Else
        Exit Sub
    End If
    
End Sub

Public Sub GetUpdateSMLW(strGetUpdate As String, intSMLW As Integer)
    Dim fInsert As Boolean
        
    Call Utility_Search("SMLW", "", "NAME", True)
    
    If gudtUtility(0).UtilityName = "" Then
        strGetUpdate = "NEW"
    End If
        
    If strGetUpdate = "GET" Then
        gintSMLW = CInt(gudtUtility(0).UtilityValueLong)
        Exit Sub
    ElseIf strGetUpdate = "UPDATE" Or strGetUpdate = "NEW" Then
    
        OpenUtilityDatabase     ' ltas 8.29.2006
        Call ErrorCheck(dbUtility)
        
        If dbUtility = 0 Then
            MsgBox "Unable to open Utility Setup database"
            Exit Sub
        End If
        
        PDBSetSortFields dbUtility, 0
        PDBMoveFirst dbUtility
        
        If strGetUpdate = "UPDATE" Then
            PDBFindRecordByField dbUtility, 0, gudtUtility(0).UtilityID
            PDBEditRecord dbUtility
            
            gudtUtility(0).UtilityValueText = CStr(intSMLW)
            gudtUtility(0).UtilityValueLong = CLng(intSMLW)
            
        Else 'NEW ' First Time ADD - Default to ON
            gudtUtility(0).UtilityName = "SMLW"
            gudtUtility(0).UtilityValueLong = 0
            gudtUtility(0).UtilityValueText = "0"
            gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
            PDBCreateRecordBySchema (dbUtility)
        End If
        
        fInsert = WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
                
        PDBClose dbUtility
        dbUtility = 0
        gintSMLW = CInt(gudtUtility(0).UtilityValueLong)

    Else
        Exit Sub
    End If
    
End Sub

Public Function Upgrade20060622(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim aryPdLine() As tPdLineRecord
    
intError = 1
    Call Utility_Search("UPGRADE20060622", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20060622" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20060622")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update Load Structures.  Adding Load Memory and MillIDs"
        lbl.Visible = True
        lbl.Refresh
        'Update Load Table
intError = 4
        Call LoadLoadPDBData
intError = 5
        FSO.DeleteFile (gstrPDBPath & "\Load.pdb")
intError = 6
               
        
        CreateDatabasePDB dbLoad, "Load", Load_Schema
        PDBClose dbLoad
        dbLoad = 0
intError = 7
        For I = 0 To UBound(garyLoad)
            garyLoad(I).MillID = 0
            garyLoad(I).LoadMem = ""
        Next
intError = 8
        SaveLoadPDBData
        
        'Update CTLine Table
        lbl.Caption = "Update Load Structures in CTLine.  Adding MillID and Time Tracking"
        lbl.Visible = True
        lbl.Refresh
intError = 9
        Call LoadgaryCTLine
intError = 10
        FSO.DeleteFile (gstrPDBPath & "\CTLine.pdb")

        CreateDatabasePDB dbCTLine, "CTLine", CTLine_Schema
        PDBClose dbCTLine
        dbCTLine = 0
intError = 11
        For I = 0 To UBound(garyCTLine)
            garyCTLine(I).AutoMatch = GetCTLineAutoMatch(garyCTLine(I))
            If IsDate(garyCTLine(I).Surface) = True Then
                garyCTLine(I).GDT = CDate(garyCTLine(I).Surface)
                garyCTLine(I).MillID = 0
                garyCTLine(I).Pieces = 1
            End If
        Next
        Call SaveCTLinePDBData
intError = 12

'Update pd Table
        lbl.Caption = "Update Load Structures in Production.  Adding MillID and Time Tracking"
        lbl.Visible = True
        lbl.Refresh
intError = 13
        Call LoadgaryPD("", "ALL")
        
intError = 14
        FSO.DeleteFile (gstrPDBPath & "\pd.pdb")

        CreateDatabasePDB dbPD, "pd", PD_Schema
        PDBClose dbPD
        dbPD = 0
intError = 15
        For I = 0 To UBound(garyPD)
            garyPD(I).PDDTStart = garyPD(I).PDRoughGradingDate
            garyPD(I).PDDTEnd = garyPD(I).PDRoughGradingDate
            garyPD(I).MillID = 0
        Next
        
        
        
        Call OpenPDDatabase
        For I = 0 To UBound(garyPD)
            PDBCreateRecordBySchema dbPD
            WritePDRecord garyPD(I)
            PDBUpdateRecord dbPD
        Next
        
        ClosePDDatabase
        dbPD = 0
        
intError = 16
        
        
'Update pd Table
        lbl.Caption = "Update Load Structures in Production Line.  Adding MillID and Time Tracking"
        lbl.Visible = True
        lbl.Refresh
intError = 17
    
        Call OpenPdLineDatabase
        PDBMoveFirst dbPDLine
        ReDim aryPdLine(0)
        I = 0
        
        Do Until PDBEOF(dbPDLine)
            If aryPdLine(0).PDID = 0 And UBound(aryPdLine) = 0 Then
                'do nothing
            Else
                I = I + 1
                ReDim Preserve aryPdLine(I)
            End If
            
            Call ReadPdLineRecord(aryPdLine(I))
            PDBMoveNext dbPDLine
        Loop
        
        ClosePdLineDatabase
        
                
intError = 18
        FSO.DeleteFile (gstrPDBPath & "\pdline.pdb")

        CreateDatabasePDB dbPDLine, "pdline", PdLine_Schema
        ClosePdLineDatabase
        
intError = 19
        
        Call OpenPdLineDatabase
        For I = 0 To UBound(aryPdLine)
            aryPdLine(I).MillID = 0
            PDBCreateRecordBySchema dbPDLine
            WritePdLineRecord aryPdLine(I)
            PDBUpdateRecord dbPDLine
        Next
        
        ClosePdLineDatabase
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        If gudtUtility(0).UtilityID = 0 Then gudtUtility(0).UtilityID = 1
        gudtUtility(0).UtilityName = "UPGRADE20060622"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
'        PDBFindRecordByField dbUtility, 1, gudtUtility(0).UtilityName
        
        PDBCreateRecordBySchema dbUtility
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        MsgBox "Upgrade Complete.  Load Memory and Mill Tracking Added.  Please restart program."
        lbl.Visible = False
        End
        
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060622 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20060622 = False
    Exit Function
    
End Function
Public Function Upgrade20060628(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim aryPdLine() As tPdLineRecord
    
intError = 1
    Call Utility_Search("UPGRADE20060628", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20060628" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20060628")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update Load Structures.  Adding Load Memory and MillIDs"
        lbl.Visible = True
        lbl.Refresh
        'Update Load Table
intError = 4
        Call LoadLoadPDBData
intError = 5
        FSO.DeleteFile (gstrPDBPath & "\Load.pdb")
intError = 6
               
        CreateDatabasePDB dbLoad, "Load", Load_Schema
        PDBClose dbLoad
        dbLoad = 0
intError = 7
        For I = 0 To UBound(garyLoad)
            garyLoad(I).TotalFootage = 0
            garyLoad(I).TotalFootageGrade = 0
            garyLoad(I).TargetPercent = 0
        Next
intError = 8
        SaveLoadPDBData
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20060628"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        Call CompactgaryCTLineData(lbl)
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Load Memory Process Upgrade Added.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060628 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20060628 = False
    Exit Function
    
End Function
Public Function Upgrade20060720(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim aryPdLine() As tPdLineRecord
    
intError = 1
    Call Utility_Search("UPGRADE20060720", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20060720" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20060720")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update CTLine Add ShiftID Structures and Reporting."
        lbl.Visible = True
        lbl.Refresh
        'Update Load Table
intError = 4
        'Update CTLine Table
intError = 9
        Call LoadgaryCTLine
intError = 10
        FSO.DeleteFile (gstrPDBPath & "\CTLine.pdb")

        CreateDatabasePDB dbCTLine, "CTLine", CTLine_Schema
        PDBClose dbCTLine
        dbCTLine = 0
intError = 11
        For I = 0 To UBound(garyCTLine)
            garyCTLine(I).ShiftID = Month(garyCTLine(I).GDT) & day(garyCTLine(I).GDT) & Year(garyCTLine(I).GDT) & "-" & garyCTLine(I).Grader
        Next
        Call SaveCTLinePDBData

'Update pd Table
        lbl.Caption = "Update Shift Structures in Production."
        lbl.Visible = True
        lbl.Refresh
intError = 13
        Call LoadgaryPD("", "ALL")
        
intError = 14
        FSO.DeleteFile (gstrPDBPath & "\pd.pdb")

        CreateDatabasePDB dbPD, "pd", PD_Schema
        PDBClose dbPD
        dbPD = 0
intError = 15
        
        Call OpenPDDatabase
        For I = 0 To UBound(garyPD)
            PDBCreateRecordBySchema dbPD
            WritePDRecord garyPD(I)
            PDBUpdateRecord dbPD
        Next
        
        ClosePDDatabase
        dbPD = 0
        
intError = 16

'Update pdline Table
        lbl.Caption = "Update Shift Structures in Production Line."
        lbl.Visible = True
        lbl.Refresh
intError = 17
    
        Call OpenPdLineDatabase
        PDBMoveFirst dbPDLine
        ReDim aryPdLine(0)
        I = 0
        
        Do Until PDBEOF(dbPDLine)
            If aryPdLine(0).PDID = 0 And UBound(aryPdLine) = 0 Then
                'do nothing
            Else
                I = I + 1
                ReDim Preserve aryPdLine(I)
            End If
            
            Call ReadPdLineRecord(aryPdLine(I))
            PDBMoveNext dbPDLine
        Loop
        
        ClosePdLineDatabase
        
                
intError = 18
        FSO.DeleteFile (gstrPDBPath & "\pdline.pdb")

        CreateDatabasePDB dbPDLine, "pdline", PdLine_Schema
        ClosePdLineDatabase
        
intError = 19
        
        Call OpenPdLineDatabase
        For I = 0 To UBound(aryPdLine)
            aryPdLine(I).ShiftID = ""
            PDBCreateRecordBySchema dbPDLine
            WritePdLineRecord aryPdLine(I)
            PDBUpdateRecord dbPDLine
        Next
        
        ClosePdLineDatabase

intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20060720"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  CT Line Upgrade Complete.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060720 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20060720 = False
    Exit Function
    
End Function

Public Function Upgrade20060726(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim aryGrade() As tGradeRecord
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20060726", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20060726" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20060726")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update Grade Structures.  Adding Cant Length Functions"
        lbl.Visible = True
        lbl.Refresh
        'Update Load Table
        OpenGradeDatabase
        
                
        ReDim aryGrade(0)
        PDBMoveFirst dbGrade
        Do Until PDBEOF(dbGrade)
            If aryGrade(0).GradeID = 0 And UBound(aryGrade) = 0 Then
                ReadGradeRecord aryGrade(0)
            Else
                ReDim Preserve aryGrade(UBound(aryGrade) + 1)
                ReadGradeRecord aryGrade(UBound(aryGrade))
            End If
            aryGrade(UBound(aryGrade)).Cant = 0
            PDBMoveNext dbGrade
        Loop
            
        CloseGradeDatabase
        
intError = 5
        FSO.DeleteFile (gstrPDBPath & "\grade.pdb")
intError = 6
               
        CreateDatabasePDB dbLoad, "grade", Grade_Schema
        PDBClose dbGrade
        dbGrade = 0
intError = 7
        
        
intError = 8
        'Save the Grade Data after structure update
        OpenGradeDatabase
    
        For I = 0 To UBound(aryGrade)
            PDBCreateRecordBySchema dbGrade
            fInsert = WriteGradeRecord(aryGrade(I))
            PDBUpdateRecord dbGrade
        Next
        
        CloseGradeDatabase
        
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20060726"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        Call CompactgaryCTLineData(lbl)
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Grade/Cant Expansion Process Upgrade Added.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060726 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20060726 = False
    Exit Function
    
End Function
Public Function Upgrade20060728(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20060728", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20060728" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20060728")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update Load and Load Memory Structures.  Adding Cant Length Functions"
        lbl.Visible = True
        lbl.Refresh
        
        'Update Load Table
intError = 4
        Call LoadLoadPDBData
intError = 5
        FSO.DeleteFile (gstrPDBPath & "\Load.pdb")
intError = 6
               
        CreateDatabasePDB dbLoad, "Load", Load_Schema
        PDBClose dbLoad
        dbLoad = 0
intError = 7
        For I = 0 To UBound(garyLoad)
            garyLoad(I).FootageTargetLoad = 0
        Next
intError = 8
        SaveLoadPDBData
        
        
        'Update LoadMemory Table
intError = 4
        Call LoadLoadMemPDBData
intError = 5
        FSO.DeleteFile (gstrPDBPath & "\LoadMem.pdb")
intError = 6
               
        CreateDatabasePDB dbLoadMem, "LoadMem", LoadMem_Schema
        PDBClose dbLoadMem
        dbLoadMem = 0
intError = 7
        For I = 0 To UBound(garyLoadMem)
            garyLoadMem(I).FootageTargetLoad = 0
        Next
intError = 8
        SaveLoadMemPDBData
        
        
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20060728"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        Call CompactgaryCTLineData(lbl)
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Load/LoadMemory Expansion Process Upgrade Added.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060728 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20060728 = False
    Exit Function
    
End Function
Public Function CheckforPDUpdate20090328() As Boolean
    Dim udt As tPDRecord
    Dim dbTemp As Long
    Dim I As Integer
    Dim intError As Integer
On Error GoTo ErrorHandler
    
    Exit Function
    
    If App.Major = 10 And App.Minor = 7 Then
        'Continue
        OpenPDDatabase
        ReDim garyPDV1(0)
intError = 1
        If PDBNumRecords(dbPD) = 0 Then
            Call DeleteDatabasePDB("pd", dbPD)
            CreateDatabasePDB dbPD, "pd", PD_Schema
            ClosePDDatabase
            dbPD = 0
            Exit Function
        End If

        If PDBGetNumFields(dbPD) <= 83 Then
            ReDim garyPDV1(PDBNumRecords(dbPD) - 1)
            'it needs update
            
intError = 5
            If PDBNumRecords(dbPD) = 0 Then
            
            Else
                PDBBulkRead dbPD, PDBNumRecords(dbPD), VarPtr(garyPDV1(0))
            End If
            ClosePDDatabase
            dbPD = 0
            
intError = 10
            'CreateDatabasePDB dbTemp, "tmp", PD_Schema
            'dbTemp = PDBOpen(Byfilename, gstrPDBPath & "\tmp", 0, 0, 0, 0, afModeReadWrite)
            
            Call DeleteDatabasePDB("pd", dbPD)
            CreateDatabasePDB dbPD, "pd", PD_Schema
            OpenPDDatabase
'            dbTemp = PDBOpen(Byfilename, gstrPDBPath & "\tmp", 0, 0, 0, 0, afModeReadWrite)
intError = 15

            For I = 0 To UBound(garyPDV1)
                If garyPDV1(I).BundleID <> "" And garyPDV1(I).BundleID <> "0" Then
                    udt.PDID = garyPDV1(I).PDID
                    udt.BundleID = garyPDV1(I).BundleID
                    udt.ThicknessID = garyPDV1(I).ThicknessID
                    udt.thickness = garyPDV1(I).thickness
                    udt.SpeciesID = garyPDV1(I).SpeciesID
                    udt.Species = garyPDV1(I).Species
                    udt.PDLayers = garyPDV1(I).PDLayers
                    udt.PDWeight = garyPDV1(I).PDWeight
                    udt.PDEstimated = garyPDV1(I).PDEstimated
                    udt.PDLocation = garyPDV1(I).PDLocation
                    udt.PDRoughGradingDate = garyPDV1(I).PDRoughGradingDate
                    udt.PDEnterKilnDate = garyPDV1(I).PDEnterKilnDate
                    udt.PDToFinalGradingDate = garyPDV1(I).PDToFinalGradingDate
                    udt.PDFinalGradeDate = garyPDV1(I).PDFinalGradeDate
                    udt.PDShipmentDate = garyPDV1(I).PDShipmentDate
                    udt.BOLID = garyPDV1(I).BOLID
                    udt.StatusID = garyPDV1(I).StatusID
                    udt.Status = garyPDV1(I).Status
                    udt.PDTotalGrossBFM = garyPDV1(I).PDTotalGrossBFM
                    udt.PDTotalNetBFM = garyPDV1(I).PDTotalNetBFM
                    udt.PDTotalPieces = garyPDV1(I).PDTotalPieces
                    udt.PDNotes = garyPDV1(I).PDNotes
                    udt.PDInspector = garyPDV1(I).PDInspector
                    udt.LengthID = garyPDV1(I).LengthID
                    udt.Length = garyPDV1(I).Length
                    udt.PDSurfaceType = garyPDV1(I).PDSurfaceType
                    udt.PDLoadNumber = garyPDV1(I).PDLoadNumber
                    udt.PDPackNumber = garyPDV1(I).PDPackNumber
                    udt.OrgID = garyPDV1(I).OrgID
                    udt.Org = garyPDV1(I).Org
                    udt.KilnNumber = garyPDV1(I).KilnNumber
                    udt.PDFinalGradeStatus = garyPDV1(I).PDFinalGradeStatus
                    udt.PDFinalGradingType = garyPDV1(I).PDFinalGradingType
                    udt.PDFinalGradingSort = garyPDV1(I).PDFinalGradingSort
                    udt.GID = garyPDV1(I).GID
                    udt.GDT = garyPDV1(I).GDT
                    udt.PDHHExport = garyPDV1(I).PDHHExport
                    udt.PDHHKilnExport = garyPDV1(I).PDHHKilnExport
                    udt.PDHHCombine = garyPDV1(I).PDHHCombine
                    udt.PDGRADES = garyPDV1(I).PDGRADES
                    udt.PDGRADESTEXT = garyPDV1(I).PDGRADESTEXT
                    udt.PDClass = garyPDV1(I).PDClass
                    udt.PDClassID = garyPDV1(I).PDClassID
                    udt.PDBatchDate = garyPDV1(I).PDBatchDate
                    udt.PDBatchID = garyPDV1(I).PDBatchID
                    udt.PDBatchDate_DT = garyPDV1(I).PDBatchDate_DT
                    udt.PDHHGradeBreakdown = garyPDV1(I).PDHHGradeBreakdown
                    udt.PDInventory = garyPDV1(I).PDInventory
                    udt.PDInventoryFlag = garyPDV1(I).PDInventoryFlag
                    udt.PDI1 = garyPDV1(I).PDI1
                    udt.PDI2 = garyPDV1(I).PDI2
                    udt.GreenDry = garyPDV1(I).GreenDry
                    udt.VendorOrg = garyPDV1(I).VendorOrg
                    udt.VendorOrgAID = garyPDV1(I).VendorOrgAID
                    udt.VendorOrgID = garyPDV1(I).VendorOrgID
                    udt.PDThicknessText = garyPDV1(I).PDThicknessText
                    udt.PDThickness = garyPDV1(I).PDThickness
                    udt.MillID = garyPDV1(I).MillID
                    udt.PDDTStart = garyPDV1(I).PDDTStart
                    udt.PDDTEnd = garyPDV1(I).PDDTEnd
                    udt.ShiftID = garyPDV1(I).ShiftID
                    udt.PDID_New = garyPDV1(I).PDID_New
                    udt.Printed = garyPDV1(I).Printed
                    udt.PONumber = garyPDV1(I).PONumber
                    udt.TotalBundleCount = garyPDV1(I).TotalBundleCount
                    udt.LoadNotes = garyPDV1(I).LoadNotes
                    udt.LoadInstruct = garyPDV1(I).LoadInstruct
                    udt.TargetGradeFootage = garyPDV1(I).TargetGradeFootage
                    udt.ActualGradeFootage = garyPDV1(I).ActualGradeFootage
                    udt.TallyDetail = garyPDV1(I).TallyDetail
                    udt.Position = garyPDV1(I).Position
                    udt.PDW1 = garyPDV1(I).PDW1
                    udt.PDW2 = garyPDV1(I).PDW2
                    udt.TallyType = garyPDV1(I).TallyType
                    udt.LRID = garyPDV1(I).LRID
                    udt.LRAID = garyPDV1(I).LRAID
                    udt.AvgWidth = garyPDV1(I).AvgWidth
                    udt.PDReceiveExport = garyPDV1(I).PDReceiveExport
                    udt.ProdRunAID = ""
                    udt.PDX1 = ""
                    udt.PDX2 = ""
                    udt.PDX3 = ""
                    udt.PDX4 = ""
                    udt.PDX5 = ""
                    
                    'CleargudtPD
                    PDBCreateRecordBySchema dbPD
                    PDBWriteRecord dbPD, VarPtr(udt)
                    PDBUpdateRecord dbPD
                End If
            Next
intError = 16
            PDBClose dbPD
            PDBClose dbPD
            dbPD = 0
        Else
            'it's done
        End If
        
        PDBClose (dbTemp)
        ClosePDDatabase
        dbPD = 0
    
        
    
    Else
        Exit Function
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "CheckforPDUp: " & intError & " (" & Err.Number & Err.Description & ")"
    Exit Function
    
End Function
Public Function Upgrade20060821(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20060821", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20060821" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\UPGRADE20060821")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update PDID_New, HHSerialNum"
        lbl.Visible = True
        lbl.Refresh
        
'Update pd Table
        lbl.Caption = "Update PD Table."
        lbl.Visible = True
        lbl.Refresh
intError = 13
        Call LoadgaryPD("", "ALL")
        
intError = 14
        FSO.DeleteFile (gstrPDBPath & "\pd.pdb")

        CreateDatabasePDB dbPD, "pd", PD_Schema
        PDBClose dbPD
        dbPD = 0
intError = 15
        
        Call OpenPDDatabase
        For I = 0 To UBound(garyPD)
            garyPD(I).PDID_New = garyPD(I).PDID
            
            PDBCreateRecordBySchema dbPD
            WritePDRecord garyPD(I)
            PDBUpdateRecord dbPD
        Next
        
        ClosePDDatabase
        dbPD = 0
        
intError = 16
        
        'Update CTLine Table
        lbl.Caption = "Update Load Structures in CTLine.  Adding MillID and Time Tracking"
        lbl.Visible = True
        lbl.Refresh
intError = 9
        Call LoadgaryCTLine
intError = 10
        FSO.DeleteFile (gstrPDBPath & "\CTLine.pdb")

        CreateDatabasePDB dbCTLine, "CTLine", CTLine_Schema
        PDBClose dbCTLine
        dbCTLine = 0
intError = 11
        For I = 0 To UBound(garyCTLine)
            garyCTLine(I).HHSerialNum = gstrHHPSerialNum
        Next
        Call SaveCTLinePDBData
intError = 12

        
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20060821"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        Call CompactgaryCTLineData(lbl)
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Load/LoadMemory Expansion Process Upgrade Added.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060728 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20060821 = False
    Exit Function
    
End Function

Public Function Upgrade20060912(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim aryPdLine() As tPdLineRecord
    
intError = 1
    Call Utility_Search("UPGRADE20060912", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20060912" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20060912")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update CTLine Add Length and Width Structures."
        lbl.Visible = True
        lbl.Refresh
        'Update Load Table
intError = 4
        'Update CTLine Table
intError = 9
        Call LoadgaryCTLine
intError = 10
        FSO.DeleteFile (gstrPDBPath & "\CTLine.pdb")

        CreateDatabasePDB dbCTLine, "CTLine", CTLine_Schema
        PDBClose dbCTLine
        dbCTLine = 0
intError = 11
        Call SaveCTLinePDBData

intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20060912"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  CT Line Upgrade Complete.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060912 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20060912 = False
    Exit Function
End Function

Public Function Upgrade20061009(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim aryPdLine() As tPdLineRecord
    
intError = 1
    Call Utility_Search("UPGRADE20061009", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20061009" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20061009")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update CTLine Add Length and Width Structures."
        lbl.Visible = True
        lbl.Refresh
        'Update Load Table
intError = 4
        'Update CTLine Table
intError = 9
        Call LoadgaryPD("", "ALL")
        
intError = 14
        FSO.DeleteFile (gstrPDBPath & "\pd.pdb")

        CreateDatabasePDB dbPD, "pd", PD_Schema
        PDBClose dbPD
        dbPD = 0
intError = 15
        
        Call OpenPDDatabase
        For I = 0 To UBound(garyPD)
            garyPD(I).PDID_New = garyPD(I).PDID
            
            PDBCreateRecordBySchema dbPD
            WritePDRecord garyPD(I)
            PDBUpdateRecord dbPD
        Next
        
        ClosePDDatabase
        dbPD = 0
        

intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20061009"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  PD Upgrade Complete.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20061009 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20061009 = False
    Exit Function
End Function

Public Function Upgrade20061105(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim aryPdLine() As tPdLineRecord
    
intError = 1
    Call Utility_Search("Upgrade20061105", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "Upgrade20061105" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\Upgrade20061105")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update TallyKey Structures."
        lbl.Visible = True
        lbl.Refresh
        'Update Load Table
intError = 4
        'Update CTLine Table
intError = 9
        ReDim gTallyKey(TallyKey.Last, 0)
        Call UpdateTallyKeyList(36)
        Call LoadTallyKey
        
intError = 14
        FSO.DeleteFile (gstrPDBPath & "\TallyKey.pdb")

        CreateDatabasePDB dbPD, "TallyKey", TallyKey_Schema
        PDBClose dbTallyKey
intError = 15
        
       SaveTallyKey

intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "Upgrade20061105"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  TallyKey Upgrade Complete.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20061105 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20061105 = False
    Exit Function
End Function


'Old Code - Delete after 8/1/06
Public Function CompactgaryCTLineData(lbl As AFLabel) As Boolean


Exit Function

    Dim aryCom() As tCTLineRecord
    Dim I As Long, J As Integer
    Dim intCom As Long
    Dim fNew As Boolean
    Dim K As Integer

    Call LoadgaryCTLine
    ReDim aryCom(0)
    aryCom(0).CTLineID = -1

    For I = 0 To UBound(garyCTLine)
        lbl.Caption = "Compressing " & UBound(garyCTLine) & " Records ... Please Wait"
        lbl.Refresh



        fNew = True
        For J = 0 To intCom
            Debug.Print "Start: " & I & ": " & garyCTLine(I).Grade & "  SM: " & garyCTLine(I).SM

            If garyCTLine(I).GradeID = aryCom(J).GradeID And _
                    garyCTLine(I).LoadAID = aryCom(J).LoadAID And _
                    garyCTLine(I).SM = aryCom(J).SM And _
                    garyCTLine(I).SpeciesID = aryCom(J).SpeciesID And _
                    garyCTLine(I).StatusID = aryCom(J).StatusID And _
                    garyCTLine(I).LengthID = aryCom(J).LengthID And _
                    garyCTLine(I).MillID = aryCom(J).MillID And _
                    garyCTLine(I).OrgID = aryCom(J).OrgID And _
                    garyCTLine(I).ThicknessID = aryCom(J).ThicknessID And _
                    garyCTLine(I).VendorID = aryCom(J).VendorID And _
                    garyCTLine(I).Surface = aryCom(J).Surface And _
                    garyCTLine(I).Grader = aryCom(J).Grader Then
                fNew = False
                Exit For
            End If
        Next

        If fNew = False Then
            aryCom(J).Pieces = aryCom(J).Pieces + garyCTLine(I).Pieces

            If IsGreen(aryCom(J).StatusID) Then
                aryCom(J).Gross = fgetFootage(CDbl(1), CDbl(aryCom(J).SM), aryCom(J).Pieces, aryCom(J).thickness) * 12
                aryCom(J).Net = fgetFootage(CDbl(1), CDbl(aryCom(J).SM), aryCom(J).Pieces, aryCom(J).thickness) * 12 * dblshrinkage

            Else    'Kiln Dried
                aryCom(J).Net = fgetFootage(CDbl(1), CDbl(aryCom(J).SM), aryCom(J).Pieces, aryCom(J).thickness) * 12
                aryCom(J).Gross = fgetFootage(CDbl(1), CDbl(aryCom(J).SM), aryCom(J).Pieces, aryCom(J).thickness) * 12 / dblshrinkage
            End If

            Debug.Print "Update: " & aryCom(J).Grade & " - " & aryCom(J).Pieces & "  SM: " & aryCom(J).SM
        ElseIf fNew = True Then
            If I = 0 And aryCom(0).CTLineID = -1 Then
                'do nothing
            Else
                intCom = intCom + 1
                ReDim Preserve aryCom(intCom)
            End If

            aryCom(intCom).CTLineID = garyCTLine(I).CTLineID
            aryCom(intCom).GDT = garyCTLine(I).GDT
            aryCom(intCom).Grade = garyCTLine(I).Grade
            aryCom(intCom).GradeID = garyCTLine(I).GradeID
            aryCom(intCom).Grader = garyCTLine(I).Grader
            aryCom(intCom).Gross = garyCTLine(I).Gross
            aryCom(intCom).Length = garyCTLine(I).Length
            aryCom(intCom).LengthID = garyCTLine(I).LengthID
            aryCom(intCom).LoadAID = garyCTLine(I).LoadAID
            aryCom(intCom).Location = garyCTLine(I).Location
            aryCom(intCom).MillID = garyCTLine(I).MillID
            aryCom(intCom).Net = garyCTLine(I).Net
            aryCom(intCom).Org = garyCTLine(I).Org
            aryCom(intCom).OrgAID = garyCTLine(I).OrgAID
            aryCom(intCom).OrgID = garyCTLine(I).OrgID
            aryCom(intCom).PDBatchID = garyCTLine(I).PDBatchID
            aryCom(intCom).SM = garyCTLine(I).SM
            aryCom(intCom).Species = garyCTLine(I).Species
            aryCom(intCom).SpeciesID = garyCTLine(I).SpeciesID
            aryCom(intCom).Status = garyCTLine(I).Status
            aryCom(intCom).StatusID = garyCTLine(I).StatusID
            aryCom(intCom).Surface = garyCTLine(I).Surface
            aryCom(intCom).thickness = garyCTLine(I).thickness
            aryCom(intCom).ThicknessID = garyCTLine(I).ThicknessID
            aryCom(intCom).Vendor = garyCTLine(I).Vendor
            aryCom(intCom).VendorAID = garyCTLine(I).VendorAID
            aryCom(intCom).VendorID = garyCTLine(I).VendorID
            aryCom(intCom).Pieces = garyCTLine(I).Pieces
            Debug.Print "NEW: " & aryCom(intCom).Grade & " " & aryCom(intCom).Pieces & "  SM: " & aryCom(intCom).SM
        End If
    Next


    ReDim garyCTLine(UBound(aryCom))

    Call LoadLoadPDBData
    For K = 0 To UBound(garyLoad)
        garyLoad(K).TotalFootage = 0
        garyLoad(K).TotalFootageGrade = 0
    Next

    For I = 0 To UBound(aryCom)
        garyCTLine(I).CTLineID = aryCom(I).CTLineID
        garyCTLine(I).GDT = aryCom(I).GDT
        garyCTLine(I).Grade = aryCom(I).Grade
        garyCTLine(I).GradeID = aryCom(I).GradeID
        garyCTLine(I).Grader = aryCom(I).Grader
        garyCTLine(I).Gross = aryCom(I).Gross
        garyCTLine(I).Length = aryCom(I).Length
        garyCTLine(I).LengthID = aryCom(I).LengthID
        garyCTLine(I).LoadAID = aryCom(I).LoadAID
        garyCTLine(I).Location = aryCom(I).Location
        garyCTLine(I).MillID = aryCom(I).MillID
        garyCTLine(I).PONumber = aryCom(I).PONumber
        garyCTLine(I).Net = aryCom(I).Net
        garyCTLine(I).Org = aryCom(I).Org
        garyCTLine(I).OrgAID = aryCom(I).OrgAID
        garyCTLine(I).OrgID = aryCom(I).OrgID
        garyCTLine(I).PDBatchID = aryCom(I).PDBatchID
        garyCTLine(I).SM = aryCom(I).SM
        garyCTLine(I).Species = aryCom(I).Species
        garyCTLine(I).SpeciesID = aryCom(I).SpeciesID
        garyCTLine(I).Status = aryCom(I).Status
        garyCTLine(I).StatusID = aryCom(I).StatusID
        garyCTLine(I).Surface = aryCom(I).Surface
        garyCTLine(I).thickness = aryCom(I).thickness
        garyCTLine(I).ThicknessID = aryCom(I).ThicknessID
        garyCTLine(I).Vendor = aryCom(I).Vendor
        garyCTLine(I).VendorAID = aryCom(I).VendorAID
        garyCTLine(I).VendorID = aryCom(I).VendorID
        garyCTLine(I).Pieces = aryCom(I).Pieces

        For K = 0 To UBound(garyLoad)
            If garyLoad(K).LoadAID = garyCTLine(I).LoadAID Then
                If IsGreen(garyLoad(K).StatusID) Then
                    garyLoad(K).TotalFootage = garyLoad(K).TotalFootage + garyCTLine(I).Gross
                    If garyLoad(K).GradeID = garyCTLine(I).GradeID Then
                        garyLoad(K).TotalFootageGrade = garyLoad(K).TotalFootageGrade + garyCTLine(I).Gross
                    End If
                Else
                    garyLoad(K).TotalFootage = garyLoad(K).TotalFootage + garyCTLine(I).Net
                    If garyLoad(K).GradeID = garyCTLine(I).GradeID Then
                        garyLoad(K).TotalFootageGrade = garyLoad(K).TotalFootageGrade + garyCTLine(I).Net
                    End If
                End If
            End If
        Next
        lbl.Caption = "Writing New Compressed Records ... " & I & " of " & UBound(aryCom)
        lbl.Refresh
    Next

    ReDim aryCom(0)
    lbl.Caption = "Saving Records..."
    lbl.Refresh
    DeleteCTLinePDBData
    SaveCTLinePDBData
    Call DeleteLoadPDBData
    Call SaveLoadPDBData

    lbl.Caption = "Compression Complete!"
    lbl.Refresh
    lbl.Visible = False
    CompactgaryCTLineData = True
End Function

Public Function Upgrade20080907(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20080907", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20080907" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
              
        Call Utility_Search("SERIALNUM", gstrHHPSerialNum, "NAME", True, True, True)
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20061110"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
                
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Load/LoadMemory Expansion Process Upgrade Added.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060728 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20080907 = False
    Exit Function
    
End Function

Public Function Upgrade20171024(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20171024", "", "NAME", True)

    If gudtUtility(0).UtilityName = "UPGRADE20171024" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2


        Call Utility_Search("SERIALNUM", gstrHHPSerialNum, "NAME", True, True, True)

        'Create the Prod and ProdType PDB's if they are not already there
On Error GoTo SkipProdCreate
        CreateDatabasePDB dbProd, "Prod", Prod_Schema
        CloseProdDatabase
SkipProdCreate:
On Error GoTo SkipProdTypeCreate
        CreateDatabasePDB dbProdType, "ProdType", ProdType_Schema
        CloseProdTypeDatabase
SkipProdTypeCreate:

intError = 20
        'Complete Update
        ReDim gudtUtility(0)

        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20171024"
        gudtUtility(0).UtilityValueText = "COMPLETE"


        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility

        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25



        gintBackupCountCT = 1
        Call BackuptoSD(lbl)

        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Prod/ProdType Expansion Completed"
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20171024 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20171024 = False
    Exit Function
    
End Function
Public Function Upgrade20171214(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20171214E", "", "NAME", True)

    If gudtUtility(0).UtilityName = "UPGRADE20171214E" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2


        Call Utility_Search("SERIALNUM", gstrHHPSerialNum, "NAME", True, True, True)

        'Create the Event and EventType PDB's if they are not already there
On Error GoTo SkipEventCreate
        CreateDatabasePDB dbEvent, "Event", Event_Schema
        CloseEventDatabase
SkipEventCreate:


intError = 20
        'Complete Update
        ReDim gudtUtility(0)

        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20171214E"
        gudtUtility(0).UtilityValueText = "COMPLETE"


        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility

        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25



        gintBackupCountCT = 1
        Call BackuptoSD(lbl)

        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Event/Event Log Expansion Completed. mLIMBS Will Close.  Please Restart the program and the new features will be enabled!"
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20171214 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20171214 = False
    Exit Function
    
End Function
Public Function Upgrade20171215(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20171215I", "", "NAME", True)

    If gudtUtility(0).UtilityName = "UPGRADE20171215I" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2


        Call Utility_Search("SERIALNUM", gstrHHPSerialNum, "NAME", True, True, True)

        'Create the Pay and PayType PDB's if they are not already there
On Error GoTo SkipPayCreate
        CreateDatabasePDB dbPay, "Pay", Pay_Schema
        ClosePayDatabase
SkipPayCreate:


intError = 20
        'Complete Update
        ReDim gudtUtility(0)

        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20171215I"
        gudtUtility(0).UtilityValueText = "COMPLETE"


        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility

        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25



        gintBackupCountCT = 1
        Call BackuptoSD(lbl)

        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Pay/Pay Log Expansion Completed. mLIMBS Will Close.  Please Restart the program and the new features will be enabled!"
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20171215 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20171215 = False
    Exit Function
    
End Function
Public Function CheckforGradeUpdate20160812() As Boolean
    Dim udt As tGradeRecord
    
    Dim I As Integer
    Dim intError As Integer
    Dim aryGradeV2() As tGradeV2Record
    
On Error GoTo ErrorHandler
    
    
    
    If App.Major <= 19 Then
        'Continue
        OpenGradeDatabase
        ReDim gudtGrade(0)
        
intError = 1
        If PDBNumRecords(dbGrade) = 0 Then
            Call DeleteDatabasePDB("grade", dbGrade)
            CreateDatabasePDB dbGrade, "grade", Grade_Schema
            CloseGradeDatabase
            dbGrade = 0
            Exit Function
        End If

        If PDBGetNumFields(dbGrade) <= 15 Then
            ReDim aryGradeV2(PDBNumRecords(dbGrade) - 1)
            'it needs update
            
intError = 5
            If PDBNumRecords(dbGrade) = 0 Then
            
            Else
                PDBBulkRead dbGrade, PDBNumRecords(dbGrade), VarPtr(aryGradeV2(0))
                
            End If
            CloseGradeDatabase
            dbGrade = 0
            
intError = 10
          
            Call DeleteDatabasePDB("grade", dbGrade)
            CreateDatabasePDB dbPD, "grade", Grade_Schema
            OpenGradeDatabase

intError = 15

            For I = 0 To UBound(aryGradeV2)
                
                udt.GradeID = aryGradeV2(I).GradeID
                udt.Grade = aryGradeV2(I).Grade
                udt.Grouping = aryGradeV2(I).Grouping
                udt.HHAID = aryGradeV2(I).HHAID
                udt.GradingUse = aryGradeV2(I).GradingUse
                udt.DisplayOrder = aryGradeV2(I).DisplayOrder
                udt.Cant = aryGradeV2(I).Cant
                udt.Width = aryGradeV2(I).Width
                udt.thickness = aryGradeV2(I).thickness
                udt.CantMask = aryGradeV2(I).CantMask
                udt.ReceivingUse = aryGradeV2(I).ReceivingUse

                udt.GradeType = ""
                udt.HuskyGradeID = ""
                udt.SilvaTechGradeID = ""
                udt.HHSortOrder = aryGradeV2(I).DisplayOrder
                udt.SOAID = ""
                udt.Protected = 0
                udt.HHActive = 1
                udt.GradeGroupFlag = 0
                udt.GGID = aryGradeV2(I).GradeID
                udt.HHAIDNum = 0
                udt.GradeX1 = ""
                udt.GradeX2 = ""
                udt.GradeX3 = ""
                udt.GradeX4 = ""
                udt.GradeX5 = ""

                    
                PDBCreateRecordBySchema dbGrade
                PDBWriteRecord dbGrade, VarPtr(udt)
                PDBUpdateRecord dbGrade
            Next
intError = 16
            PDBClose dbGrade
            dbPD = 0
            MsgBox "V.19.x Grade Table Update Completed!"
        Else
            'it's done
        End If
        
        PDBClose (dbGrade)
        CloseGradeDatabase
        dbGrade = 0
        
    Else
        Exit Function
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "CheckforGradeTableUpdate: " & intError & " (" & Err.Number & Err.Description & ")"
    Exit Function
    
End Function


Public Function Upgrade20061110(lbl As AFLabel) As Boolean
    Dim FSO As CFileManager
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Call Utility_Search("UPGRADE20061110", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20061110" Then
        'Upgrade already completed
        Exit Function
    Else
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\UPGRADE20061110")
intError = 3
        Set FSO = New CFileManager
        

        lbl.Caption = "Update PDID_New, HHSerialNum"
        lbl.Visible = True
        lbl.Refresh
        
'Update pd Table
        lbl.Caption = "Update PD Table."
        lbl.Visible = True
        lbl.Refresh
intError = 13
        Call LoadgaryPD("", "ALL")
        
intError = 14
        FSO.DeleteFile (gstrPDBPath & "\pd.pdb")

        CreateDatabasePDB dbPD, "pd", PD_Schema
        PDBClose dbPD
        dbPD = 0
intError = 15
        
        Call OpenPDDatabase
        For I = 0 To UBound(garyPD)
            garyPD(I).PDID_New = garyPD(I).PDID
            
            PDBCreateRecordBySchema dbPD
            WritePDRecord garyPD(I)
            PDBUpdateRecord dbPD
        Next
        
        ClosePDDatabase
        dbPD = 0
        
intError = 16
        
              
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20061110"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        MsgBox "Upgrade Complete.  Load/LoadMemory Expansion Process Upgrade Added.  Please restart program."
        End
    End If

    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20060728 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20061110 = False
    Exit Function
    
End Function
Public Function Upgrade20080318(lbl As AFLabel) As Boolean
    Dim I As Integer
    Dim intError As Integer
    Dim fInsert As Boolean

intError = 1
    Upgrade20080318 = True
    
    Call Utility_Search("UPGRADE20080318", "", "NAME", True)
        
    If gudtUtility(0).UtilityName = "UPGRADE20080318" Then
        'Upgrade already completed
        Exit Function
    Else
        Upgrade20080318 = False
intError = 2
        
        MsgBox "Please Wait, a system Upgrade Is Processing, you will be notified when Completed."
      
        Call BackuptoSD(lbl, True, "\UPGRADE20080318")
intError = 3
        

        lbl.Caption = "Creating LR Table"
        lbl.Visible = True
        lbl.Refresh
        CreateDatabasePDB dbLR, "LR", LR_Schema
        PDBClose dbLR
        dbLR = 0
        
        lbl.Caption = "Creating LRLine Table"
        lbl.Visible = True
        lbl.Refresh
        
        CreateDatabasePDB dbLRLine, "LRLine", LRLine_Schema
        PDBClose dbLRLine
        dbLRLine = 0
    
        
        
        
intError = 20
        'Complete Update
        ReDim gudtUtility(0)
        
        gudtUtility(0).UtilityID = GetNextIDPDB(dbUtility, 0)
        gudtUtility(0).UtilityName = "UPGRADE20080318"
        gudtUtility(0).UtilityValueText = "COMPLETE"
        
        OpenUtilityDatabase
        PDBCreateRecordBySchema dbUtility
        
        Call WriteUtilityRecord(gudtUtility(0))
        PDBUpdateRecord dbUtility
        CloseUtilityDatabase
intError = 25

        gintBackupCountCT = 1
        Call BackuptoSD(lbl)
        
        lbl.Caption = "UPDATE COMPLETE"
        lbl.Refresh
        lbl.Visible = False
        Upgrade20080318 = True
        MsgBox "Upgrade Complete.  Load Receiving Update Complete.  Please restart program."
        End
    End If

    
    
    Exit Function
ErrorHandler:
    MsgBox "Error Stop: 00" & intError & " - 20080318 has failed.  Please Contact eLIMBS Tech Support at 888.520.1951. PLEASE RECORD ERROR STOP ID NOW!"
    Upgrade20080318 = False
    Exit Function
    
End Function


Public Sub cleansdcard(lbl As AFLabel)
    Dim FSO As CFileManager
    Dim strStorageCardPath As String
    Dim fSDBackup As Boolean

On Error GoTo ErrorHandler
    Set FSO = New CFileManager
    Dim fDir As CDirectory
    Dim tmpfil As cFile
    Dim strTemp As String
    Dim lngTemp As Long
    Dim I As Integer
    
    strStorageCardPath = "\storage card"


    #If AppForge Then 'continue
    #Else
        Exit Sub
    #End If

    lbl.Caption = "Checking for SD"
'
    Set fDir = FSO.OpenDirectory(strStorageCardPath, False)
    strTemp = fDir.EnumFirstDirectory(lngTemp)
    
    I = 0
    Do Until strTemp = ""
        I = I + 1
        strTemp = fDir.EnumNextDirectory(lngTemp)
        lbl.Caption = "Removing #" & I & " strtemp"
        lbl.Refresh
        FSO.DeleteDirectory strTemp
    Loop

    lbl.Caption = "Clean SD Card Completed!"
    Set FSO = Nothing

    Exit Sub

ErrorHandler:
    lbl.Caption = "Error During SD Card Cleaning - " & Err.Number & Err.Description & ""
    Exit Sub

End Sub
Public Sub LoadPDBArray()
    Dim I As Integer

On Error GoTo ErrorHandler

    ReDim garyPDB(34)
    I = -1
    
    I = I + 1
    garyPDB(I) = "LR.pdb"
    I = I + 1
    garyPDB(I) = "ProdRun.pdb"
    I = I + 1
    garyPDB(I) = "LRLine.pdb"
    I = I + 1
    garyPDB(I) = "LoadMem.pdb"
    I = I + 1
    garyPDB(I) = "Utility.pdb"
    I = I + 1
    garyPDB(I) = "BE.pdb"
    I = I + 1
    garyPDB(I) = "CTLine.pdb"
    I = I + 1
    garyPDB(I) = "grade.pdb"
    I = I + 1
    garyPDB(I) = "length.pdb"
    I = I + 1
    garyPDB(I) = "Load.pdb"
    I = I + 1
    garyPDB(I) = "org.pdb"
    I = I + 1
    garyPDB(I) = "productionstatus.pdb"
    I = I + 1
    garyPDB(I) = "species.pdb"
    I = I + 1
    garyPDB(I) = "TallyKey.pdb"
    I = I + 1
    garyPDB(I) = "thickness.pdb"
    I = I + 1
    garyPDB(I) = "user.pdb"
    I = I + 1
    garyPDB(I) = "pd.pdb"

    I = I + 1
    garyPDB(I) = "pdline.pdb"
    I = I + 1
    garyPDB(I) = "OH.pdb"
    I = I + 1
    garyPDB(I) = "OL.pdb"
    I = I + 1
    garyPDB(I) = "TAGINV.pdb"
    I = I + 1
    garyPDB(I) = "MAC.pdb"
    I = I + 1
    garyPDB(I) = "LOC.pdb"
    I = I + 1
    garyPDB(I) = "PA.pdb"
    I = I + 1
    garyPDB(I) = "CONT.pdb"
    I = I + 1
    garyPDB(I) = "RH.pdb"
    I = I + 1
    garyPDB(I) = "RL.pdb"
    I = I + 1
    garyPDB(I) = "GradeDist.pdb"
    I = I + 1
    garyPDB(I) = "ProdType.pdb"
    I = I + 1
    garyPDB(I) = "Prod.pdb"
    
    I = I + 1
    garyPDB(I) = "Event.pdb"
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in LoadPDBArray " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub
Public Function Upgrade10_9_1Check() As Boolean
 Dim FSO As CFileManager
 Dim intError As Integer

On Error GoTo ErrorHandler

 Set FSO = New CFileManager
intError = 1
    
    dbCTLine = PDBOpen(Byfilename, "\Program Files\ITToolworks\mLIMBS\CTLine", 0, 0, 0, 0, afModeReadWrite)
 intError = 2
 
    If dbCTLine <> 0 Then
        Upgrade10_9_1Check = True
    Else
       ClosePDDatabase
       dbPD = 0
       Upgrade10_9_1Check = False
    End If
intError = 5
    Exit Function
ErrorHandler:
    MsgBox intError & " - Position Error during upgrade check."
    Exit Function
End Function
Public Sub Upgrade10_9_1RUN()
    Dim FSO As CFileManager
    Dim intError As Integer
    Dim intErrorCount As Integer
    Dim I As Integer
    Dim fBackup As Boolean
    
intError = 1
    
    closealldatabases
    Call LoadPDBArray
    
    Set FSO = New CFileManager
                
    If BackuptoSD(frmBackup.lblBackup, True, "\Storage Card\Upgrade10_9_1", True) = False Then
        MsgBox "Contact Technical Support PreUpgrade Backup Failed"
        End
        Exit Sub
    End If
    
    Call FSO.OpenDirectory("\IPSM\mLIMBS", True)
    
intError = 2
    For I = 0 To UBound(garyPDB)
        intError = intError + 1
        frmUserQuestion.lblQuestion.Caption = "Moving File: " & garyPDB(I)
        FSO.MoveFile "\Program Files\ITToolworks\mLIMBS\" & garyPDB(I), "\IPSM\mLimbs\" & garyPDB(I)
    Next
    
    Set FSO = Nothing
    frmUserQuestion.lblQuestion.Caption = "Upgrade Complete"
    Unload frmUserQuestion
    frmLogin.Show
    
    Exit Sub

ErrorHandler:
    If intError = 1 Then
        MsgBox "Contact Technical Support PreUpgrade Backup Failed"
        End
        Exit Sub
    Else
        MsgBox "Contact Technical Support - Upgrade Error (" & intError & " - " & garyPDB(I) & " Failed to Copy"
        End
        Exit Sub
    End If
    
End Sub

Public Function ZebraBarCodePrintEstimated(address As String, port As Long, strfile As String, BTSerial As AFSerial, Optional socketWebData As AFClientSocket) As Boolean
    
    ZebraBarCodePrintEstimated = True
    Dim msg As String
    Dim sIPAddress As String
    Dim tmp As String
    Dim FSO As CFileManager
    Dim str As CFileTextReadable
    Dim intError As Integer
    Dim fDebug As Boolean
    
    Set FSO = New CFileManager

intError = 1
On Error GoTo ErrorHandler
           
    'fDebug = True
    '
    If fDebug = True Then
        MsgBox "Debug Mode = TRUE - Web Printing"
        address = "192.168.21.147"
        sIPAddress = address
        gPrinterPort = 6101
        
        
        strfile = "EndTally.lbl"
            
            socketWebData.Close
            If (sIPAddress = "") Then
                socketWebData.LocalPort = TimerMS And &H7FFF
                sIPAddress = socketWebData.ResolveHostName(address)
            End If
            socketWebData.Protocol = afSocketProtocolTCP
            socketWebData.RemoteHostIP = sIPAddress
            socketWebData.RemotePort = gPrinterPort
            socketWebData.Connect sIPAddress, gPrinterPort, 5 * 1000           ' If it cannot connect to the printer this command will hang the program for 5 seconds.
    Else
    
        If port = 1 Then
            BTSerial.Settings = "9600,N,8,1"
        Else
            BTSerial.Settings = "57600,N,8,1"
        End If
    End If

    If fDebug = True And address <> "" Then
        Dim strTmp As String
        Dim strresponse As String
        socketWebData.Close
        
        socketWebData.Protocol = afSocketProtocolTCP
        socketWebData.LocalPort = TimerMS And &H7FFF
        socketWebData.RemoteHostIP = gPrinterIP
        socketWebData.RemotePort = gPrinterPort
        socketWebData.Connect gPrinterIP, gPrinterPort, 500
        If socketWebData.LastError = -2 Then
            socketWebData.SendString chr(6), 1, 20
            socketWebData.SendString chr(2) & "AB", 3, 20
            socketWebData.SendString chr(6), 1, 20
        Else
            strTmp = " "
            Do While strTmp <> ""
            socketWebData.GetString strTmp, 1, 1000
            strresponse = strresponse & strTmp
            Loop
            MsgBox ("Connected to Printer. Ready to Print Tags.")
        End If
        
    Else
    
        intError = 2
            BTSerial.PortOpen = False
        intError = 3
            BTSerial.CommPort = port
        intError = 4
        #If AppForge Then
            BTSerial.PortOpen = True
        #Else
          'uncomment for pc direct print via BT  BTSerial.PortOpen = True
        #End If
        
        intError = 5
        'MsgBox BTSerial.Settings
        'MsgBox address <> ""
    End If

Set str = FSO.OpenReadOnlyAsText(gstrPDBPath & "\" & strfile)
    
    tmp = "BEGIN"
    
    Do While tmp <> ""
        tmp = str.ReadLine
        Debug.Print tmp
       
        If tmp = "" Or tmp = "<END>" Then Exit Do
       
        If gudtPD.PDRoughGradingDate > #1/1/2000# Then
           tmp = Replace(tmp, "<DATE>", FormatDateTime(gudtPD.PDRoughGradingDate, vbShortDate))
        Else
            tmp = Replace(tmp, "<DATE>", "")
        End If
        
        tmp = Replace(tmp, "<LOAD>", gudtPD.PDLoadNumber)
        
        
        tmp = Replace(tmp, "<BUNDLEID>", gudtPD.BundleID)
        tmp = Replace(tmp, "<BUNDLE>", gudtPD.BundleID)
        tmp = Replace(tmp, "<12345678>", gudtPD.BundleID)
        
        tmp = Replace(tmp, "<RUNID>", tcu(gudtPD.ProdRunAID))
        tmp = Replace(tmp, "<RUN#>", tcu(gudtPD.ProdRunAID))
        
        tmp = Replace(tmp, "<B->", Replace(gudtPD.BundleID, "70G", ""))
        
        tmp = Replace(tmp, "<LOCATION>", gudtPD.PDLocation)
        If gudtPD.PDPackNumber <> 0 Then
            tmp = Replace(tmp, "<PACK>", gudtPD.PDPackNumber)
        Else
            tmp = Replace(tmp, "<PACK>", "")
        End If
        
        tmp = Replace(tmp, "<STATUS>", GetStatusData(CStr(gudtPD.StatusID), "ID", "PSAID"))
        tmp = Replace(tmp, "<SPECIES>", GetSpeciesData(CStr(gudtPD.SpeciesID), "ID", "HHAID"))
        tmp = Replace(tmp, "<SPECIESNAME>", GetSpeciesData(CStr(gudtPD.SpeciesID), "ID", "SPECIES"))
        
        tmp = Replace(tmp, "<GRADENAME>", GetGradeData(gudtPD.PDGRADES, "ID", "GRADE"))
        tmp = Replace(tmp, "<GRADEAID>", GetGradeData(gudtPD.PDGRADES, "ID", "HHAID"))
        tmp = Replace(tmp, "<GRADE>", GetGradeData(gudtPD.PDGRADES, "ID", "HHAID"))
        
        tmp = Replace(tmp, "<COLORAID>", CStr(gudtPD.ColorAID))
        tmp = Replace(tmp, "<COLOR>", gudtPD.ColorAID)
        tmp = Replace(tmp, "<COLORDESC>", CStr(GetPAData(gudtPD.ColorAID, "HHAID", "COLOR", True, "PADESC")))
        tmp = Replace(tmp, "<COLORNAME>", GetPAData(gudtPD.ColorAID, "HHAID", "COLOR", True, "PADESC"))
        
        tmp = Replace(tmp, "<WIDTHAID>", gudtPD.PDW1)
        tmp = Replace(tmp, "<WIDTH>", gudtPD.PDW1)
        tmp = Replace(tmp, "<WIDTHNAME>", GetPAData(gudtPD.PDW1, "HHAID", "WIDTH", True, "PADESC"))
        tmp = Replace(tmp, "<WIDTHDESC>", GetPAData(gudtPD.PDW1, "HHAID", "WIDTH", True, "PADESC"))
        
        tmp = Replace(tmp, "<PDCLASS>", gudtPD.PDClass)
        tmp = Replace(tmp, "<CLASSAID>", gudtPD.PDClass)
        tmp = Replace(tmp, "<CLASS>", gudtPD.PDClass)
        
        tmp = Replace(tmp, "<SURFACE>", gudtPD.PDSurfaceType)
        tmp = Replace(tmp, "<ORDERID>", gudtPD.PONumber)
        
        tmp = Replace(tmp, "<PDI1>", gudtPD.PDI1)
        tmp = Replace(tmp, "<PDI2>", gudtPD.PDI2)
                
        tmp = Replace(tmp, "<PDINSPECTOR>", gudtPD.PDInspector)
        tmp = Replace(tmp, "<PSINSP>", gudtPD.PDInspector)
        tmp = Replace(tmp, "<INSP>", gudtPD.PDInspector)
        
        tmp = Replace(tmp, "<CUSTOMER>", gudtPD.Org)
        tmp = Replace(tmp, "<VENDOR>", gudtPD.VendorOrg)
        
        tmp = Replace(tmp, "<CUSTOMERNAME>", GetOrgData(CStr(gudtPD.OrgID), "ID", "", "NAME"))
        tmp = Replace(tmp, "<VENDORNAME>", GetOrgData(CStr(gudtPD.VendorOrgID), "ID", "", "NAME"))
        
        
        
        If gudtPD.LengthID <= 0 Then
            tmp = Replace(tmp, "<LENGTHS>", "")
            tmp = Replace(tmp, "<LENGTHNAME>", "")
            tmp = Replace(tmp, "<LENGTH>", "")
        Else
            tmp = Replace(tmp, "<LENGTHS>", GetLengthData(CStr(gudtPD.LengthID), "ID", "LENGTHNAME"))
            tmp = Replace(tmp, "<LENGTHNAME>", GetLengthData(CStr(gudtPD.LengthID), "ID", "LENGTHNAME"))
            tmp = Replace(tmp, "<LENGTH>", GetLengthData(CStr(gudtPD.LengthID), "ID", "LENGTHNAME"))
        End If
        
        
        'Below handhelds the cant type grades or dimensioned bundle types or lumber bundle types depending on the tallytype field and grade/cant flag
        
        
        If SC(gSettings.BTStaves, "YES") = True Then
            tmp = Replace(tmp, "<THK>", gudtPD.thickness)
            tmp = Replace(tmp, "<PCS>", gudtPD.PDTotalPieces)
            tmp = Replace(tmp, "<AVGWIDTH>", CStr(Round(CDbl(gudtPD.PDW1) / CDbl(gudtPD.PDTotalPieces), 2)))
            
        ElseIf GetGradeData(gudtPD.PDGRADES, "HHAID", "CANT") = "1" Or SC(gudtPD.TallyType, "DIMENSIONED") = True _
            Or SC(gudtPD.TallyType, "CNDIMENSIONED") = True And SC(gudtPD.TallyType, "DIMENSIONEDBASIC") = True Or InStr(tcu(gudtPD.TallyType), "DIMENSION") > 0 Then
            'For Dimensioned Products
            If GetGradeData(gudtPD.PDGRADES, "HHAID", "CANT") = "1" Then
                tmp = Replace(tmp, "<THK>", "")
            Else
                tmp = Replace(tmp, "<THK>", gudtPD.thickness)
                tmp = Replace(tmp, "Layers", "Pieces")
            End If
            
            If gudtPD.PDTotalPieces <> 0 Then
                tmp = Replace(tmp, "<Pieces>", gudtPD.PDTotalPieces)
            ElseIf gudtPD.PDLayers <> 0 Then
                tmp = Replace(tmp, "<Pieces>", gudtPD.PDLayers)
            Else
                tmp = Replace(tmp, "<Pieces>", "")
                tmp = Replace(tmp, "PCS:", "")
                tmp = Replace(tmp, "Pieces:", "")
            End If
            'End of Dimensioned Products
        Else
            'Standard Lumber Products - Not Dimensioned
            tmp = Replace(tmp, "<THK>", GetThicknessData(CStr(gudtPD.ThicknessID), "ID", "THICKNESS"))
            If gudtPD.PDLayers <> 0 Then
                tmp = Replace(tmp, "<LAYERS>", gudtPD.PDLayers)
            Else
                tmp = Replace(tmp, "<LAYERS>", "")
                tmp = Replace(tmp, "Layers:", "")
            End If
            'end non-dimensioned
        End If
        If gudtPD.PDTotalPieces <> 0 And (InStr(tcu(tmp), "PCS") > 0 Or InStr(tcu(tmp), "PIECES") > 0) Then
            tmp = Replace(tmp, "<Pieces>", gudtPD.PDTotalPieces)
            tmp = Replace(tmp, "<PCS>", gudtPD.PDTotalPieces)
            tmp = Replace(tmp, "<PIECES>", gudtPD.PDTotalPieces)
        End If
        
        'For Multi Grade Estimated Bundles with line grades listed
        'For each grade in the bundle, this will list the individual line grades and the volume
        
        If SC(gudtPD.TallyType, "INVENTORY") = True Then
            'No PDLine Records for Inventory Lookup bundles
        Else
            If SC(gudtPD.TallyType, "CNESTMULTI") = True Or SC(gudtPD.TallyType, "BLKESTMULTI") = True Or SC(gudtPD.TallyType, "BLKESTMULTIBASIC") = True Then
                Dim K As Integer
                For K = 0 To UBound(gudtPDLine)
                    If K <= 6 Then
                        tmp = Replace(tmp, "<L-GRADE" & CStr(K + 1) & ">", GetGradeData(CStr(gudtPDLine(K).GradeID), "ID", "HHAID"))
                        tmp = Replace(tmp, "<L-GROSS" & CStr(K + 1) & ">", CStr(gudtPDLine(K).PDLineGross))
                        tmp = Replace(tmp, "<L-NET" & CStr(K + 1) & ">", CStr(gudtPDLine(K).PDLineNet))
                    End If
                Next
            Else
                tmp = Replace(tmp, "<L-GRADE1>", GetGradeData(CStr(gudtPDLine(0).GradeID), "ID", "HHAID"))
                tmp = Replace(tmp, "<L-GROSS1>", CStr(gudtPDLine(0).PDLineGross))
                tmp = Replace(tmp, "<L-NET1>", CStr(gudtPDLine(0).PDLineNet))
            End If
                
        'Now get rid of any remaining line grade tags that weren't already matched
            For K = 0 To 6
                If SC(Left(tmp, 1), "B") = True Then
                    If K > UBound(gudtPDLine) And SC(Right(Trim(tmp), 2), "-" & K + 1) = True Then
                        tmp = "XXX"
                    End If
                End If
                                
                tmp = Replace(tmp, "<L-GRADE" & CStr(K + 1) & ">", "")
                tmp = Replace(tmp, "<L-GROSS" & CStr(K + 1) & ">", "")
                tmp = Replace(tmp, "<L-NET" & CStr(K + 1) & ">", "")
            Next
        End If
        
        'Volume / Footage Field Options
        tmp = Replace(tmp, "<GROSSFOOTAGE>", Round(gudtPD.PDTotalGrossBFM, 0))
        tmp = Replace(tmp, "<NETFOOTAGE>", Round(gudtPD.PDTotalNetBFM, 0))
                
        If IsGreen(gudtPD.StatusID) = True Then
            If gudtPD.PDTotalGrossBFM <> 0 Then
                tmp = Replace(tmp, "<FOOTAGE>", Round(gudtPD.PDTotalGrossBFM, 0))
            Else
                tmp = Replace(tmp, "<FOOTAGE>", "")
            End If
        Else
            If gudtPD.PDTotalNetBFM <> 0 Then
                tmp = Replace(tmp, "<FOOTAGE>", Round(gudtPD.PDTotalNetBFM, 0))
            Else
                tmp = Replace(tmp, "<FOOTAGE>", "")
            End If
        End If
        
        
'''        If chkDipped.Value = 1 Then
'''            tmp = Replace(tmp, "<X1>", "X")
'''        Else
'''            tmp = Replace(tmp, "<X1>", "")
'''        End If
'''
'''        If chkEndSealed.Value = 1 Then
'''            tmp = Replace(tmp, "<X2>", "X")
'''        Else
'''            tmp = Replace(tmp, "<X2>", "")
'''        End If
        tmp = replaceTags(tmp, gudtPD)
        tmp = Replace(tmp, chr(13), "")
        tmp = Replace(tmp, chr(10), "")
        If tmp <> "XXX" Then
            msg = tmp & chr(13) & chr(10)
        
       'SocketWebData.SendString msg, Len(msg), 1 * 1000
intError = 7

            If address <> "" Then
                  socketWebData.SendString msg, Len(msg), 200
            Else
                   BTSerial.Output = msg
            End If
            
        End If
        Debug.Print tmp
    Loop
'    MsgBox ("finished printing")
intError = 6
Debug.Print msg

    Set FSO = Nothing
    Set str = Nothing
    
    '''socketWebData.Close
    If address <> "" Then
        '''SocketWebData.Disconnect
        
        socketWebData.Close
    Else
        BTSerial.PortOpen = False
    End If
intError = 8
    Exit Function
ErrorHandler:
    ZebraBarCodePrintEstimated = False
    MsgBox "Failed To Print: " & Err.Number & "-" & Err.Description & " - Address: " & address & " Port " & port & " File " & strfile & "(" & intError & ")  " & Err.Number & Err.Description
    
    Exit Function
    
End Function

Public Function ZebraTagPrint(strBundleID As String, ByRef BTSerial As AFSerial, _
    Optional webSocketData As AFClientSocket, Optional strReturnTagFileName As String, _
    Optional strPrintType As String) As Boolean
    'Case "CNESTMULTI", "CNDIMENSIONED", "BLKESTMULTIBASIC", "BLKESTMULTI", "DIMENSIONEDBASIC", "BUNDLETALLY", "BTESTIMATED", "CNESTIMATED", "DIMENSIONED", "LWCHAINTALLY", "BLKESTIMATED", "BLKESTIMATEDBASIC"


On Error GoTo ErrorHandler
    ZebraTagPrint = False
    
    'this is just to keep it the way it was/backward compatible with the default being btprinttype
    If SC(strPrintType, "") = True Then
        strPrintType = gSettings.BTPrintType
    End If
    
    If SCInList(strPrintType, "NETWORKFILE,NETWORKFILEV2,NETWORKFILEV3,V3,REPLACETAGS") = True Then
        If PrintTag_AnyByBundleID(gudtPD.PDID, gudtPD.BundleID, "", BTSerial, webSocketData, True, False, "", "", True) = True Then
            MsgBox "Bundle " & gudtPD.BundleID & "(" & gudtPD.TallyType & ")  Network File/Print Completed Successfully!"
        Else
            MsgBox ("Bundle Failed to Print(" & gSettings.BTPrintType & "). Please Check the printer and make sure that it is active and try again.")
            Exit Function
        End If
    ElseIf SC(gSettings.ZPL_Print_Language, "ZPL") = True Then
        Call func_ZebraBundleBarCodePrint_ZPL(CLng(gSettings.BTComm), gSettings.ZPL_Tag_File_Name, BTSerial)
        ZebraTagPrint = True
        
    ElseIf SC(gudtPD.TallyDetail, "CNDIMENSIONED") = True Or SC(gudtPD.TallyType, "DIMENSIONEDBASIC") = True Or _
        SC(gudtPD.TallyDetail, "BLKESTIMATED") = True Or SC(gudtPD.TallyType, "BLKESTIMATEDBASIC") Or _
        Trim(UCase(gudtPD.TallyType)) = "CNESTIMATED" Or Trim(UCase(gudtPD.TallyType)) = "INVENTORY" Or Trim(UCase(gudtPD.TallyType)) = "BLOCKTALLY" _
        Or Left(Trim(UCase(gudtPD.TallyType)), 11) = "DIMENSIONED" Or Trim(UCase(gudtPD.TallyType)) = "DIMENSIONED" Then
        
        strReturnTagFileName = "ChainTag.lbl"
        Call ZebraBarCodePrintEstimated(gPrinterIP, glngBTCommReceive, "ChainTag.lbl", BTSerial, webSocketData)
        
        
    ElseIf InStr(Trim(UCase(gudtPD.TallyType)), "MULTI") > 0 Then
        strReturnTagFileName = "BTTag2.lbl"
        Call ZebraBarCodePrintEstimated(gPrinterIP, glngBTCommReceive, "BTTag1.lbl", BTSerial, webSocketData)
        
        
        strReturnTagFileName = "BTTag2.lbl"
        If SC(gSettings.BTTagFileCount, "1") = False Then Call ZebraBarCodePrintEstimated(gPrinterIP, glngBTCommReceive, "BTTag2.lbl", BTSerial, webSocketData)
        
    Else
        
        Dim lngBTCommET As Long
        Dim strETTagFile As String
        
        Call LoadgudtPDLine(gudtPD.PDID, "PDID,MillID")
        If SC(gSettings.ETTagFileName, "") = True Then
            Call Utility_Search("TAGFILE-ET", "", "NAME", True, True)
            gSettings.ETTagFileName = gudtUtility(0).UtilityValueText
        End If
        
        If gSettings.BTComm <= 0 Then
            MsgBox "The End Tally Printer Comm Port is not Specified"
ZebraTagPrint = False
            Exit Function
        End If
    
        Call func_ZebraBundleBarCodePrintBT(CLng(gSettings.BTComm), gSettings.ETTagFileName, BTSerial)

ZebraTagPrint = True
        Exit Function
    End If
ZebraTagPrint = True
    Exit Function
ErrorHandler:
ZebraTagPrint = False
    MsgBox "Error in Print to Zebra: " & Err.Number & "-" & Err.Description & vbCrLf & "Tag File: " & strReturnTagFileName
    
    Exit Function
End Function
Function FixStringWidth(strString As String, intWidth As Integer, Optional fInFront As Boolean, Optional fZeroFillInFront As Boolean) As String
    Dim intDiff As Integer
    Dim strPad As String
    Dim I As Integer
    
    If Len(strString) < intWidth Then
        intDiff = intWidth - Len(strString)
        For I = 1 To intDiff
            strPad = strPad & " "
        Next
    Else
        strString = Left(strString, intWidth)
        strPad = ""
    End If
    
    If fInFront = True Then
        If fZeroFillInFront = True Then
            strPad = Replace(strPad, " ", "0")
        End If
        strString = strPad & strString
    Else
        strString = strString & strPad
    End If
    
    FixStringWidth = strString
    
End Function
Public Function ExportWoodScanFile(fleWrite As CFileStream, lngPDID As Long) As String
    Dim strLine As String
    Dim intRow As Integer
    Dim I As Integer
    
    CleargudtPD
    CleargudtPDLine
    Call LoadgudtPD("", gudtPD, "PDID", lngPDID)
    Call LoadgudtPDLine(lngPDID, "", "", False)
    
    intRow = 0
    
    For I = 0 To UBound(gudtPDLine)
    
        'Transaction Code, Woodscan Location, Order Group, Order #
        strLine = "A" & FixStringWidth(Left(Trim(gudtPD.PDLocation), 3), 3) & FixStringWidth(Left(Trim(gudtPD.PDI1), 2), 2) & FixStringWidth(Left(Trim(gudtPD.PDLoadNumber), 8), 8)
        
        'Ticket#, Date: 00YYMMDD
        strLine = strLine & FixStringWidth(Left(Trim(gudtPD.BundleID), 8), 8) & "00" & FixStringWidth(Year(gudtPD.PDRoughGradingDate), 2, True, True) & FixStringWidth(Month(gudtPD.PDRoughGradingDate), 2, True, True) & FixStringWidth(day(gudtPD.PDRoughGradingDate), 2, True, True)
        
        'Time
        strLine = strLine & "00" & FixStringWidth(Hour(gudtPD.PDRoughGradingDate), 2, True, True) & FixStringWidth(Minute(gudtPD.PDRoughGradingDate), 2, True, True) & FixStringWidth(Second(gudtPD.PDRoughGradingDate), 2, True, True)
        
        'PIDParent
        MsgBox "PID"
        
        'Row Number, AisleNumber,
        strLine = strLine & FixStringWidth(Left(Trim(gudtPD.PDI2), 3), 3, True, False) & FixStringWidth(Left(Trim(gudtPD.PDI1), 4), 3, True, False)
        
        'TicketGroup , TicketLot
        'strLine=strLine &
    Next
    
    
End Function

Public Function CheckForDuplicateInAFGrid(grdData As AFGrid, strData As String, lblInfo As AFLabel) As Boolean
    Dim I As Integer
    Dim fFound As Boolean
    
    fFound = False
    If grdData.Rows = 0 Then Exit Function
    
    For I = 0 To grdData.Rows - 1
        If tcu(grdData.TextMatrix(I, 0)) = tcu(strData) Then
            fFound = True
            Exit For
        End If
    Next
    
    If fFound = True Then
        lblInfo.Caption = strData & " IS ALREADY IN THE LIST!"
        lblInfo.BackColor = vbYellow
        lblInfo.ForeColor = vbBlack
        lblInfo.ZOrder 0
        lblInfo.Visible = True
    Else
        lblInfo.Visible = False
    End If
    
    CheckForDuplicateInAFGrid = fFound
    
    Exit Function
ErrorHandler:
    MsgBox "Error In CheckforDuplicateInAFGrid: " & Err.Number & "-" & Err.Description
    Exit Function
End Function

Public Function IncrementAlphaNumeric(strVal As String, Optional strAlphaNumeric As String, Optional strLimit As String, Optional ByRef lngValueReturn As Long, Optional lngValueReturnOriginal As Long) As String
    Dim I As Integer
    Dim intStart As Integer
    Dim strPrefix As String
    Dim strValue As Long

On Error GoTo ErrorHandler
    If strAlphaNumeric = "ALPHA" Then
        If Len(strVal) = 0 Then Exit Function
        Dim fFoundAlpha As Boolean
        
        intStart = 0
        For I = 1 To Len(strVal)
            
            intStart = intStart + 1
            
            If IsNumeric(Mid(strVal, intStart, 1)) = True Then
                If intStart <> 1 Then
                    strValue = CLng(Mid(strVal, intStart, Len(strVal)))
                    strPrefix = Mid(strVal, 1, intStart - 1)
                End If
                Exit For
            Else
                fFoundAlpha = True
                strPrefix = Mid(strVal, 1, 1)
            End If
        Next
        
        If IsNumeric(strValue) = True And fFoundAlpha = True Then
            If strLimit = "" Then
                If Len(strVal) = 1 Then
                    IncrementAlphaNumeric = chr(Asc(strPrefix) + 1)
                Else
                    IncrementAlphaNumeric = chr(Asc(strPrefix) + 1) & (strValue)
                End If
            ElseIf Trim(strLimit) <> "" And Asc(strPrefix) + 1 > Asc(strLimit) Then
                If Len(strVal) = 1 Then
                    IncrementAlphaNumeric = chr(Asc(strPrefix) + 1)
                Else
                    IncrementAlphaNumeric = "A" & strValue
                End If
            Else
                If Len(strVal) = 1 Then
                    IncrementAlphaNumeric = chr(Asc(strPrefix) + 1)
                Else
                    IncrementAlphaNumeric = chr(Asc(strPrefix) + 1) & (strValue)
                End If
            End If
            
        Else
            IncrementAlphaNumeric = ""
        End If
    Else
        If Len(strVal) = 0 Then Exit Function
        Dim fFoundNumeric As Boolean
        
        intStart = 0
        For I = 1 To Len(strVal)
            intStart = intStart + 1
            If IsNumeric(Mid(strVal, intStart, 1)) = True And (Mid(strVal, intStart, 1)) <> "0" Then
                fFoundNumeric = True
                If intStart = 1 Then
                    strPrefix = ""
                Else
                    strPrefix = Mid(strVal, 1, intStart - 1)
                End If
                If IsNumeric((Mid(strVal, intStart, Len(strVal)))) = True Then strValue = CLng((Mid(strVal, intStart, Len(strVal))))
                
                Exit For
            End If
        Next
        
        If IsNumeric(strValue) = True And fFoundNumeric = True Then
            If IsNumeric(strLimit) = True Then
                If CLng(strLimit) <> 0 And CLng(strValue + 1) > CLng(strLimit) Then
                    IncrementAlphaNumeric = strPrefix & "1"
                    lngValueReturn = 1
                Else
                    IncrementAlphaNumeric = strPrefix & (strValue + 1)
                    lngValueReturn = strValue + 1
                End If
                
            Else
                IncrementAlphaNumeric = strPrefix & (strValue + 1)
                lngValueReturn = strValue + 1
            End If
        Else
            IncrementAlphaNumeric = ""
        End If
        lngValueReturnOriginal = lngValueReturn - 1
    End If
    
    Exit Function
ErrorHandler:
    MsgBox "Error in IncrementAlphaNumeric: " & Err.Number & "-" & Err.Description
    Exit Function
    
End Function

Public Function funcAlphaNumericAddOne(strString As String, strAlphaNumeric As String, strLimit As String) As String
    strString = UCase(strString)
    If SC(strAlphaNumeric, "NUMERIC") = True Then
        strString = IncrementAlphaNumeric(strString, "NUMERIC", strLimit)
    Else
        strString = IncrementAlphaNumeric(strString, "ALPHA", strLimit)
    End If
        
    funcAlphaNumericAddOne = strString
    
End Function

Public Sub ValidatePosition(txt As AFTextBox, strLocAID As String)
    Dim fFound As Boolean, fValid As Boolean
    Dim I As Integer
    
    fFound = False
    fValid = False
    For I = 0 To UBound(garyPA)
        If SC(garyPA(I).PAName, txt.Text) = True Then
            fFound = True
            If SC(strLocAID, garyPA(I).PATextGroup) = True Then
                fValid = True
            End If
            Exit For
        ElseIf SC(garyPA(I).PAName, tcu(strLocAID) & "-" & tcu(txt.Text)) = True Then
            fFound = True
            If SC(strLocAID, garyPA(I).PATextGroup) = True Then
                fValid = True
            End If
            Exit For
        End If
    Next
    
    If fValid = True Then
        txt.BackColor = vbWhite
    Else
        txt.BackColor = vbRed
        txt.Text = ""
    End If
    
End Sub

Public Function BundleIDFormat(strBundleID As String) As String

On Error GoTo ErrorHandler

    If gSettings.BundleIDFormat = "" Then
        BundleIDFormat = Trim(UCase(strBundleID))
    ElseIf gSettings.BundleIDFormat = "X-123" Or gSettings.BundleIDFormat = "*-???" Then
        If Len(Trim(strBundleID)) > 4 And InStr(strBundleID, "-") = 0 Then
            BundleIDFormat = Left(Trim(strBundleID), Len(Trim(strBundleID)) - 3) & "-" & Right(Trim(strBundleID), 3)
        Else
            BundleIDFormat = tcu(strBundleID)
        End If
    Else
        BundleIDFormat = Trim(UCase(strBundleID))
    End If
    Exit Function
    
ErrorHandler:
    BundleIDFormat = Trim(UCase(strBundleID))
    Exit Function
End Function

Public Function RR(strString As String) As Double
    If IsNumeric(Replace(Replace(Replace(Replace(Replace(strString, "%", ""), "$", ""), ",", ""), "'", ""), """", "")) = True Then
        RR = CDbl(Replace(Replace(Replace(Replace(Replace(strString, "%", ""), "$", ""), ",", ""), "'", ""), """", ""))
    Else
        RR = 0
    End If
    
End Function

Public Function ValidateLocationByStatus(strLocAID As String, strStatusFilter As String) As Boolean
    Dim I As Integer, J As Integer
    
    ValidateLocationByStatus = False
    
On Error Resume Next
    For I = 0 To UBound(garyLoc)
        If SC(garyLoc(I).LocAID, strLocAID) = True Then
            Call AppForge_Split(garyLoc(I).PSAID, pstrSplit, ",")
            For J = 0 To UBound(pstrSplit.ary)
                If SC(pstrSplit.ary(J), strStatusFilter) = True Then
                    ValidateLocationByStatus = True
                    Exit Function
                End If
            Next
            Exit Function
        End If
    Next
            


End Function

Public Sub WindowsRemoveShortCuts()
    Dim intError As Integer
    
On Error GoTo ErrorHandler

    If SC(gSettings.WindowsDeleteShortCuts, "YES") = True Then
        'continue
    Else
        Exit Sub
    End If
    
    #If AppForge Then
    
    #Else
        Exit Sub
        
    #End If
        
    Dim FSO As CFileManager
    
    Set FSO = New CFileManager
    
    FSO.DeleteDirectory "\Windows\Start Menu\Programs\Games"
    FSO.DeleteDirectory "\Windows\Start Menu\Programs\Office Mobile 2010"
    
    FSO.DeleteFile "\Windows\Start Menu\Programs\Alarms"
    
    FSO.DeleteFile "\Windows\Start Menu\Programs\Calendar"
    
    FSO.DeleteFile "\Windows\Start Menu\Programs\Contacts"
    FSO.DeleteFile "\Windows\Start Menu\Programs\E-mail"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Image Profiler"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Internet Explorer"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Internet Sharing"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Marketplace"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Messenger"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Modem Link"
    FSO.DeleteFile "\Windows\Start Menu\Programs\MSN Money"
    FSO.DeleteFile "\Windows\Start Menu\Programs\MSN Weather"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Notes"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Pictures & Videos"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Remote Desktop Mobile"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Search Phone"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Task Manager"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Tasks"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Text"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Windows Live"
    FSO.DeleteFile "\Windows\Start Menu\Programs\Windows Media"
    
    
    Set FSO = Nothing
    
    Exit Sub

ErrorHandler:
    Set FSO = Nothing
    
    Exit Sub
    
End Sub



Public Sub CompactgudtPDLineData()
    Dim udt() As tPdLineRecord
    Dim fmatch As Boolean
    Dim I As Integer, J As Integer
    
    ReDim udt(0)
    
    udt(0) = gudtPDLine(0)
    
    For I = 1 To UBound(gudtPDLine)
        fmatch = False
        For J = 0 To UBound(udt)
            'Check to see if there are matches that we can then compress into multi piece count rowas and minimze
            'overhead and improve speed of process of reports, entries, imports, etc
            If udt(J).PDID = gudtPDLine(I).PDID And _
                udt(J).GradeID = gudtPDLine(I).GradeID And _
                udt(J).PDLineLength = gudtPDLine(I).PDLineLength And _
                udt(J).PDLineWidth = gudtPDLine(I).PDLineWidth And _
                udt(J).ThicknessID = gudtPDLine(I).ThicknessID And _
                udt(J).ShiftID = gudtPDLine(I).ShiftID And _
                udt(J).RecAID = gudtPDLine(I).RecAID And _
                udt(J).MillID = gudtPDLine(I).MillID Then  'It is a match and can be combined by just adding the pieces
                
                udt(J).PDLinePieces = udt(J).PDLinePieces + gudtPDLine(I).PDLinePieces
                udt(J).PDLineGross = udt(J).PDLineGross + gudtPDLine(I).PDLineGross
                udt(J).PDLineNet = udt(J).PDLineNet + gudtPDLine(I).PDLineNet
                udt(J).PiecesAdj = udt(J).PiecesAdj + gudtPDLine(I).PiecesAdj

                fmatch = True
                Exit For
            End If
        Next
        If fmatch = False Then
            ReDim Preserve udt(UBound(udt) + 1)
            udt(UBound(udt)) = gudtPDLine(I)
        End If
    Next
    
    ReDim gudtPDLine(UBound(udt))
    
    For I = 0 To UBound(udt)
        gudtPDLine(I) = udt(I)
        gudtPDLine(I).ShiftID = gstrShifID
    Next
    
End Sub



Public Function PDDistrubteGradesbyGroup(lngPDID As Long, strLocAID As String, lngStatusID As Long, lngSpeciesID As Long, lngThicknessID As Long, strThickAID As String, lngGGID As Long, dblBundleFootage As Double) As Boolean
    Dim I As Integer
    Dim strMatchID As String
    Dim fisGreen As Boolean
    Dim J As Integer
    
        
    strMatchID = tcu(strLocAID) & "~" & CStr(lngStatusID) & "~" & CStr(lngSpeciesID) & "~" & CStr(lngThicknessID) & "~" & CStr(lngGGID)
    
    If GradeDistSearch(strMatchID) = True Then
        'continue
    Else
        PDDistrubteGradesbyGroup = False
        Exit Function 'since there was an error loading the grade rows, and just save it normally in main module
    End If
    
    ReDim gudtPDLine(UBound(garyGradeDist))
    fisGreen = IsGreen(lngStatusID)
    
    Dim dblTotalAdjust As Double
    dblTotalAdjust = 0
    
    For I = 0 To UBound(garyGradeDist)
        
        gudtPDLine(I).PDLineID = GetNextIDPDB(dbPDLine, 0) + I
        gudtPDLine(I).PDID = lngPDID
        gudtPDLine(I).thickness = tcu(strThickAID)
        gudtPDLine(I).ThicknessID = lngThicknessID
        gudtPDLine(I).GradeID = garyGradeDist(I).GradeID
        
        gudtPDLine(I).GDT = Now
        gudtPDLine(I).GID = glngUserID
        gudtPDLine(I).ShiftID = gstrShifID
        
        If IsNumeric(gudtPD.Length) = True Then
            gudtPDLine(I).PDLineLength = CDbl(gudtPD.Length)
        Else
            gudtPDLine(I).PDLineLength = 0
        End If
        
        If IsNumeric(gudtPD.PDW1) = True Then
            gudtPDLine(I).PDLineWidth = CDbl(gudtPD.PDW1)
        Else
            gudtPDLine(I).PDLineWidth = 0
        End If
        
        If fisGreen = True Then
            gudtPDLine(I).PDLineGross = Round(garyGradeDist(I).GradeDistPercent / 100 * dblBundleFootage, 0)
            gudtPDLine(I).PDLineNet = Round(gudtPDLine(I).PDLineGross * dblshrinkage, 0)
            dblTotalAdjust = dblTotalAdjust + gudtPDLine(I).PDLineGross
        Else
            gudtPDLine(I).PDLineNet = Round(garyGradeDist(I).GradeDistPercent / 100 * dblBundleFootage, 0)
            gudtPDLine(I).PDLineGross = Round(gudtPDLine(I).PDLineNet / dblshrinkage, 0)
            dblTotalAdjust = dblTotalAdjust + gudtPDLine(I).PDLineNet
        End If
        gudtPDLine(I).PDLineNote = "GD: L=" & gudtPD.Length & "  W=" & gudtPD.PDW1
        
        If IsNumeric(gudtPD.PDW1) = True Then
            gudtPDLine(I).PDLineWidth = CDbl(gudtPD.PDW1)
        End If
    Next
    
    'Now check the total footage we sent in, vs the total footage after the distribution of grade percentages.
    'If they don't match (due to rounding) add the extra footage to the lowest grade line (last row in gudtPDLine)
    
    If dblTotalAdjust <> dblBundleFootage Then
        If fisGreen = True Then
            gudtPDLine(UBound(gudtPDLine)).PDLineGross = gudtPDLine(UBound(gudtPDLine)).PDLineGross + (dblBundleFootage - dblTotalAdjust)
            gudtPDLine(UBound(gudtPDLine)).PDLineNet = Round(gudtPDLine(UBound(gudtPDLine)).PDLineGross * dblshrinkage, 0)
        Else
            gudtPDLine(UBound(gudtPDLine)).PDLineNet = gudtPDLine(UBound(gudtPDLine)).PDLineNet + (dblBundleFootage - dblTotalAdjust)
            gudtPDLine(UBound(gudtPDLine)).PDLineGross = Round(gudtPDLine(UBound(gudtPDLine)).PDLineNet / dblshrinkage, 0)
        End If
    End If
        
    PDDistrubteGradesbyGroup = True
    
    Exit Function
ErrorHandler:
    PDDistrubteGradesbyGroup = False
    MsgBox "Error in PDDistributeGradesByGroup " & Err.Number & "-" & Err.Description
    Exit Function
End Function

Public Function funcFileExists(strFilePathandName As String) As Boolean
    funcFileExists = False
    
    On Error GoTo ErrorHandler
    
    Dim FSO As CFileManager
    Set FSO = New CFileManager
    Dim fileTest As CFileTextReadable
    
    Set fileTest = FSO.OpenAsText(strFilePathandName, afFileModeOpen)

    Set fileTest = Nothing
    Set FSO = Nothing
    funcFileExists = True
    Exit Function
    
ErrorHandler:
    Set FSO = Nothing
    Set fileTest = Nothing
    funcFileExists = False
    Exit Function
    
End Function

Public Function PrintTag_AnyByBundleID(lngPDID As Long, strBundleID As String, _
                            Optional strPrintType As String, Optional BTSerial As AFSerial, _
                            Optional WebClientSocket As AFClientSocket, _
                            Optional fSkipLoadBundleTo_gudtPD As Boolean, _
                            Optional fSkipLoadBoardsTo_garyPDLine As Boolean, _
                            Optional ByVal strNetworkFilePath As String, _
                            Optional ByVal strNetworkFileName As String, _
                            Optional fDontClear_gudtPDWhenDone As Boolean, _
                            Optional strPrintTypeFieldList As String, _
                            Optional strReturnMessage As String, Optional fUseReturnLabel As Boolean, _
                            Optional lblReturnMessage As AFLabel) As Boolean

    
On Error GoTo ErrorHandler
    
        
    strReturnMessage = "PrintTag_AnyByBundleID In Values(PDID=" & lngPDID & ", BundleID=" & strBundleID & " "
    strReturnMessage = strReturnMessage & "PrintType=" & strPrintType & ",Path=" & strNetworkFilePath
    strReturnMessage = strReturnMessage & ", Name=" & strNetworkFileName & ", FieldList=" & strPrintTypeFieldList & ")"
    
    
    PrintTag_AnyByBundleID = False
    
    'Special Print for Stella Jones Custom frmReceiveV3 (maybe used for others in future but as of 1/28/2017 - just them so some may be hardcoded)
    If SC(strPrintType, "RECEIVEV3") = True Then
        PrintTag_AnyByBundleID = PrintTag_AnyByReceiveID(lngPDID, strBundleID, 0, strPrintType, _
                                                        strPrintTypeFieldList, strReturnMessage, BTSerial, True, _
                                                        strNetworkFilePath, strNetworkFileName, lblReturnMessage, WebClientSocket)
                                    
        Exit Function
    End If
    
    If fSkipLoadBundleTo_gudtPD = False Then ' load the bundle to the global bundle variables
        CleargudtPD
        gudtPD.PDID = -1
    
        If lngPDID > 0 Then
            Call LoadgudtPD(strBundleID, gudtPD, "PDID", lngPDID)
        ElseIf SC(strBundleID, "") = False Then
            Call LoadgudtPD(strBundleID, gudtPD, "BUNDLEID", 0)
        Else
            PrintTag_AnyByBundleID = False
            Exit Function
        End If
    End If
    
    If gudtPD.PDID > 0 Then
        If fSkipLoadBoardsTo_garyPDLine = True Then
            'do nothing should already be in there
        Else 'load boards
            Call LoadgudtPDLine(gudtPD.PDID, "", gstrShifID)
        End If
    Else
        MsgBox "Cound not Load Bundle# " & strBundleID & " OR DBID=" & lngPDID & ") "
        PrintTag_AnyByBundleID = False
        Exit Function
    End If
    
    strReturnMessage = strReturnMessage & "BundleID afer Load=" & gudtPD.BundleID
    
    Dim strTTSetType As String
    strTTSetType = ""
    If SC(strPrintType, "") = True Then
        strReturnMessage = strReturnMessage & vbCrLf & "***PrintType_Blank_NotSentInto_PrintTag_AnyByBundleID Module"
    End If
    
    'If print type isn't set at this point, use the bundle/type specific setting
    If SC(strPrintType, "") = True And SC(gudtPD.TallyType, "BUNDLETALLY") = True Then
        strPrintType = gSettings.ETPrintType
        strTTSetType = "ET"
        
    ElseIf SC(strPrintType, "") = True And InStr(tcu(gudtPD.TallyType), "CHAIN") > 0 Then
        strPrintType = tcu(gSettings.CTPrintType)
        strTTSetType = "CHAIN"
    ElseIf SC(strPrintType, "") = True And InStr(tcu(gudtPD.TallyType), "REC") > 0 Then
        strPrintType = tcu(gSettings.RecPrintType)
        strTTSetType = "REC"
    ElseIf SC(strPrintType, "") = True And _
        (InStr(tcu(gudtPD.TallyType), "EST") > 0 _
         Or InStr(tcu(gudtPD.TallyType), "BLK") > 0 _
         Or InStr(tcu(gudtPD.TallyType), "DIMENS") > 0 _
         Or InStr(tcu(gudtPD.TallyType), "MULTI") > 0) Then
    
        strPrintType = tcu(gSettings.BTPrintType)
        strTTSetType = "EST"
    ElseIf SC(strPrintType, "") = True Then
        strTTSetType = "NOMATCH-"
        strPrintType = "MOBILE"
    End If
    
    strReturnMessage = strReturnMessage & vbCrLf & "PrintType Blank-Auto Set to ** " & strTTSetType & " ** by TallyType=" & gudtPD.TallyType & vbCrLf & _
        "Before setting path/name, field list by tally type the settings are respectively: " & strNetworkFilePath & "," & strNetworkFileName & "," & strPrintTypeFieldList
    strTTSetType = ""
    
    'Now get the corresponding print type if blank, file/path if blank, and field list if blank
    If SC(gudtPD.TallyType, "BUNDLETALLY") = True Then
        If SC(strNetworkFileName, "") = True Then strNetworkFileName = gSettings.ETTagFileName
        If SC(strNetworkFilePath, "") = True Then strNetworkFilePath = gSettings.ETTagFilePath
        
        If SC(strPrintTypeFieldList, "") = True Then strPrintTypeFieldList = gSettings.ETPrintType_FieldList
        strTTSetType = "ET"
    ElseIf InStr(tcu(gudtPD.TallyType), "CHAIN") > 0 Then
        If SC(strNetworkFileName, "") = True Then strNetworkFileName = gSettings.CTTagFileName
        If SC(strNetworkFilePath, "") = True Then strNetworkFilePath = gSettings.CTTagFilePath
        If SC(strPrintTypeFieldList, "") = True Then strPrintTypeFieldList = gSettings.CTPrintType_FieldList
        strTTSetType = "CT"
    ElseIf InStr(tcu(gudtPD.TallyType), "REC") > 0 Then
        If SC(strNetworkFileName, "") = True Then strNetworkFileName = gSettings.RecTagFileName
        If SC(strNetworkFilePath, "") = True Then strNetworkFilePath = gSettings.RecTagFilePath
        If SC(strPrintTypeFieldList, "") = True Then strPrintTypeFieldList = gSettings.RecPrintType_FieldList
        strTTSetType = "REC"
    Else
        If SC(strNetworkFileName, "") = True Then strNetworkFileName = gSettings.BTTagFileName
        If SC(strNetworkFilePath, "") = True Then strNetworkFilePath = gSettings.BTTagFilePath
        If SC(strPrintTypeFieldList, "") = True Then strPrintTypeFieldList = gSettings.BTPrintType_FieldList
        strTTSetType = "BT"
    End If
    
    If SC(strTTSetType, "") = False Then strReturnMessage = strReturnMessage & vbCrLf & "Path/Name/Field Now set to=" & strNetworkFilePath & "," & strNetworkFileName & "," & strPrintTypeFieldList

    'Print Type if not set above or needs replaced
    'IF printtype is blank and etprinttype isn't us it
    If SC(strPrintType, "") = True Then
        MsgBox "Printing Not Completed" & vbCrLf & "Print Type (Mobile/Network/Wifi) Is Not Set for " & vbCrLf & _
            "TallyType=" & gudtPD.TallyType & " - Contact eLIMBS for Help @ 888.520.1951"
        PrintTag_AnyByBundleID = False
        Exit Function
    End If
    
    If SC(strPrintType, "WIFI") = True Then
        strReturnMessage = strReturnMessage & vbCrLf & "PrintType Changed from WIFI to WIRELESS"
        strPrintType = "WIRELESS" 'REplace short option with code/needed option
    End If
    
    Dim fPrintReturnedSuccess As Boolean
    fPrintReturnedSuccess = False
    
    strReturnMessage = strReturnMessage & vbCrLf & " ***Print Type Sending=" & strPrintType
    
    'If printing to a tag printer...regardless of how
    If SC(strPrintType, "MOBILE") = True Or SC(strPrintType, "WIRELESS") = True Then
        'no need to get file paths
        If SC(strPrintType, "MOBILE") = True Then
            fPrintReturnedSuccess = ZebraTagPrint(gudtPD.BundleID, BTSerial, WebClientSocket)
        ElseIf SC(strPrintType, "WIRELESS") = True = True Then
            fPrintReturnedSuccess = ZebraTagPrint(gudtPD.BundleID, BTSerial, WebClientSocket)
        End If
        strReturnMessage = strReturnMessage & vbCrLf & " ***ZEBRATAGPRINT***"
    'if writing a newtork file that another app uses to print to tag printer
    Else 'all the remaining options are network file/writing based
        If SC(strNetworkFileName, "") = True Or SC(strNetworkFilePath, "") = True Then
            MsgBox "The Network File Path or Network File Name (Setting XXTagFilePath or XXTagFileName Where XX is ET,CT,BT,REC) for Tally Type=" & gudtPD.TallyType & " Is not Configured. Printing Terminated!"
            PrintTag_AnyByBundleID = False
            Exit Function
        End If
        
        If SC(strPrintType, "NETWORKFILE") = True Then
            fPrintReturnedSuccess = func_WriteETNetworkFile(gudtPD, strNetworkFilePath, strNetworkFileName)
            strReturnMessage = strReturnMessage & vbCrLf & " ***func_WriteETNetworkFile***"
        ElseIf SC(strPrintType, "NETWORKFILEV2") = True Then
            fPrintReturnedSuccess = func_WriteETNetworkFile(gudtPD, strNetworkFilePath, strNetworkFileName, "V2")
            strReturnMessage = strReturnMessage & vbCrLf & " ***func_WriteETNetworkFile V2 Setting***"
        ElseIf SC(strPrintType, "NETWORKFILE_REPLACETAGS") = True Or SC(strPrintType, "NETWORKFILEV3") = True Or SC(strPrintType, "V3") = True Then '2016 replacetags style of network file print ..option
            fPrintReturnedSuccess = func_WriteBundleNetworkFile_ReplaceTags(gudtPD, strNetworkFilePath, strNetworkFileName, strPrintTypeFieldList, "", "", True)
            strReturnMessage = strReturnMessage & vbCrLf & " ***func_WriteBundleNetworkFile_ReplaceTags***Path=" & strNetworkFilePath & ",Name=" & strNetworkFileName & ",PrintType=" & strPrintType
        End If
    End If
    
    If fDontClear_gudtPDWhenDone = True Then
        'do nothing
    Else
        'clear the global pd and pdline variables
        CleargudtPD
        ReDim gudtPDLine(0)
    End If
    
    If SC(gSettings.DebugMode, "TAGPRINT") = True Then
        MsgBox "Tag Printing Debug Message: " & strReturnMessage
        If fUseReturnLabel = True Then
            lblReturnMessage.Caption = strReturnMessage
            lblReturnMessage.Height = 269
            lblReturnMessage.Width = 240
            lblReturnMessage.Top = 0
            lblReturnMessage.Left = 0
            lblReturnMessage.ZOrder 0
            lblReturnMessage.Visible = True
        End If
    End If
    
    PrintTag_AnyByBundleID = fPrintReturnedSuccess
    
    Exit Function
ErrorHandler:
On Error GoTo ErrorQuitAndExit
    MsgBox "Error in cmdPrintTag_Click (Print Method=" & strPrintType & "  PDID=" & lngPDID & "   B#=" & strBundleID & ") " & vbCrLf & "VB Error Message=" & Err.Number & "-" & Err.Description
    PrintTag_AnyByBundleID = False
    CleargudtPD
    ReDim gudtPDLine(0)
    Exit Function
ErrorQuitAndExit:
    PrintTag_AnyByBundleID = False
    Exit Function
End Function
Public Function PrintTag_AnyByReceiveID(lngLRID As Long, strLRLoadAID As String, intNumberOfTagsToPrint As Integer, _
                                        strPrintType As String, strPrintTypeFieldList As String, _
                                        strReturnMessage As String, _
                                        BTSerial As AFSerial, fLRRecAlreadyLoaded As Boolean, _
                                        Optional strNetworkFilePath As String, _
                                        Optional strNetworkFileName As String, _
                                        Optional lblReturnMessage As AFLabel, _
                                        Optional WebClientSocket As AFClientSocket, Optional fUseStatusLabel As Boolean, Optional lblStatus As AFLabel) As Boolean

'
    Dim udtLR As tLRRecord
On Error GoTo ErrorHandler
    
    If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID Pre-Load LR Record lngLRID=" & lngLRID & "  LoadAID=" & strLRLoadAID
    
    If fLRRecAlreadyLoaded = False Then
        Call LoadLRRecord_FromPDB(lngLRID, udtLR, strLRLoadAID)
    End If
    
    If udtLR.LRID <= 0 Then 'no record loaded/or at least not a real/valid one
        PrintTag_AnyByReceiveID = False
        If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID NOT LOADED/FOUND EXIT PRINT - LR Record lngLRID=" & lngLRID & "  LoadAID=" & strLRLoadAID
        Exit Function
    End If
    
    ReDim garyLRLine(0)
    Call LoadLRLine(udtLR.LRID, "LRID,LRLineID") ' Loads records to garyLRLine()
    
    If SC(udtLR.LengthName, "V3") = True Or SC(udtLR.LengthName, "RECEIVEV3") = True Then 'This is the field I am storing the V3 version flag for frmRecieveV3 use ...it wasn't needed for anything else
        
        If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID PREV3_PRINTTAG MODULE CALL " & UBound(garyLRLine) & " LR Line records loaded!" & vbCrLf
        If fUseStatusLabel = True Then
            lblStatus.Caption = "Preparing to Print Tags V3 "
            lblStatus.Refresh
        End If
        
        PrintTag_AnyByReceiveID = PrintTag_AnyByReceiveID_V3(udtLR, intNumberOfTagsToPrint, strReturnMessage, "", BTSerial, strNetworkFilePath, strNetworkFileName, lblReturnMessage, WebClientSocket, 0, fUseStatusLabel, lblStatus)
        
        If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID POST V3 PRINT TAG MODULE CALL " & UBound(garyLRLine) & " LR Line records loaded!" & vbCrLf
        Exit Function
    Else
        MsgBox "This version of the Receiving Print/Tag MOdule only works for Receive V3 Version Settings!"
        PrintTag_AnyByReceiveID = False
        Exit Function
    End If
    
    
    ''
    PrintTag_AnyByReceiveID = True
''*****************END CUT AND PAST FROM frmRECEIVESCAN PrintTag_AnyByReceiveID
    Exit Function
    
ErrorHandler:
On Error GoTo ErrorQuitAndExit
    MsgBox "Error in PrintTag_AnyByReceiveID (LRID=" & lngLRID & " Load# " & strLRLoadAID & " Print Method=" & strPrintType & vbCrLf & "Error# " & Err.Number & "-" & Err.Description
    PrintTag_AnyByReceiveID = False
    Exit Function

ErrorQuitAndExit:
    PrintTag_AnyByReceiveID = False
    Exit Function

End Function

Public Function PrintTag_AnyByReceiveID_V3(udtLR As tLRRecord, intNumberOfTagsToPrint As Integer, strPrintType As String, _
                                           strReturnMessage As String, BTSerial As AFSerial, _
                                            strTagFilePath As String, strTagFileName As String, _
                                           Optional lblReturnMessage As AFLabel, _
                                           Optional WebClientSocket As AFClientSocket, Optional intProdIndexOfTagToPrint As Integer, _
                                           Optional fUseStatusLabel As Boolean, Optional lblStatus As AFLabel) As Boolean
    Dim I As Long, J As Long
    Dim intError As Long
    
    Dim intCountofTotal As Integer
    Dim strLine As String
    Dim intTagCurrentRow As Integer
    Dim fFirstRow As Boolean
    Dim aryTagData() As String
    
    Dim fMixedProductsFlag As Boolean
    Dim fSkipThisProductTags As Boolean
    Dim strReturnErrorMessage As String
    Dim aryMixTags As aryStringType
    Dim intReturnNextBundleID As Integer
    Dim intReturnMixedBundle_TotalCount As Integer
    
    Dim intTotalBundleCount_Mixed_NonMixed As Integer
    
On Error GoTo ErrorHandler

intError = 0
    If PrintTag_ConnectToBTPrinter(strPrintType, BTSerial, strReturnErrorMessage) = False Then
        MsgBox "Error in PrintTag_AnyByReceiveID_V3 :  " & strReturnErrorMessage & vbCrLf & "Please be sure printer is on and try again!"
        PrintTag_AnyByReceiveID_V3 = False
        Exit Function
    End If
    
    intReturnNextBundleID = 0
    '*** Get the Mixed Tags Array to Print
    ReDim aryMixTags.ary(RecV3MixTags.Last, 0)
    If PrintTag_AnyByReceiveID_V3Mixed(udtLR, aryMixTags, intReturnNextBundleID, BTSerial, strReturnErrorMessage, strTagFilePath, gSettings.Receivev3_MixProdTag_FileName, intReturnMixedBundle_TotalCount, fUseStatusLabel, lblStatus) = False Then
        'just continue for now I guess
    End If
    
    
    '*** END OF Mixed Tag Section
    
intError = 4
    strLine = fso_FileOpen_ReadOnly_ReturnFileTextOnly("", strTagFilePath, strTagFileName)
    If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3 _ Tag File loaded to variable"
    
    'Now the full .lbl file is in the variable above (strLINE
   Debug.Print "*****TAG FILE AS READ FROM " & strTagFilePath & strTagFileName & "   ******" & vbCrLf & strLine
intError = 5
    ReDim aryTagData(0) As String
    fFirstRow = False
    
    intCountofTotal = intReturnNextBundleID
    
    intTotalBundleCount_Mixed_NonMixed = intReturnMixedBundle_TotalCount
    
    'Start temporary total count fix for nonmixed bundles/load total bundles
    Dim intTotalBundlesOnLoad As Integer
    intTotalBundlesOnLoad = intReturnMixedBundle_TotalCount
    For I = 0 To UBound(garyLRLine)
        If SC(garyLRLine(I).LRLineX1, "1") = True Then
            'ignore
        Else
            intTotalBundlesOnLoad = intTotalBundlesOnLoad + CInt(garyLRLine(I).LineBundleCount)
        End If
    Next
    
    udtLR.BundleCount = intTotalBundlesOnLoad
    
    'end temp fix
    
intError = 2000
    fSkipThisProductTags = False
    fMixedProductsFlag = False
    
    For I = 0 To UBound(garyLRLine)
        fMixedProductsFlag = False
        If SC(garyLRLine(I).LRLineX1, "1") = True Then
'''            'this is a mixed product row, only print this count once
'''            If fMixedProductsFlag = True Then
'''                'already printed skip printing
'''
'''                fSkipThisProductTags = True
'''                'print the mixed tags as identified one time
'''                garyLRLine(I).LineBundleCount = 0 'don't print them again
'''            Else
'''                fMixedProductsFlag = True
'''            End If
            fMixedProductsFlag = True 'reset every line now, since handled separately
'*** This is now handled by the module call above in the mixed tag section, so just skip them now entirely
        Else
         
            For J = 1 To garyLRLine(I).LineBundleCount
                
                If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3-Starting Tag Print Row Build garyLRLine(" & I & ") For Bundle " & J & " of " & garyLRLine(I).LineBundleCount
                Dim strTagFileToReplace As String, strTagFile_AfterReplace As String
                'Initialize variables passed to replacetags module
                strTagFileToReplace = strLine
                strTagFile_AfterReplace = ""
                
                intError = intError + (I * J)
        
                If I = 0 Or (UBound(aryTagData) = 0 And aryTagData(0) = "") Then 'redim tag data array to count in this row
                    If fFirstRow = False Then
                        ReDim aryTagData(garyLRLine(0).LineBundleCount)  ' initialize the printing array to the first lrline # of bundles...then redim preserve as wee move on to other records
                        fFirstRow = True
                        intTagCurrentRow = 0
                    Else
                        intTagCurrentRow = intTagCurrentRow + 1
                    End If
                Else
                    If J = 1 Then
                        intTagCurrentRow = UBound(aryTagData) + 1
                        ReDim Preserve aryTagData(UBound(aryTagData) + garyLRLine(J).LineBundleCount)
                    Else
                        intTagCurrentRow = intTagCurrentRow + 1
                    End If
                End If
                
                intCountofTotal = intCountofTotal + 1
                intTotalBundleCount_Mixed_NonMixed = intTotalBundleCount_Mixed_NonMixed + 1
                
                'now2 we have the positions for the data/to print, just need to fill the print array current row...through
                aryTagData(intTagCurrentRow) = strLine
    
                
                If PrintTag_AnyByReceiveID_ReplaceTags_LRPrint(udtLR, garyLRLine(I), strTagFileToReplace, strTagFile_AfterReplace, CInt(J), intCountofTotal) = True Then
                    'just keep going this module above just replaced all the tagged fields with the actual data or blanks like <LOAD> w/ R13566 orr whatever..and so on
                Else
                    MsgBox "An error occurred trying to replace field tags in the file/tag setup! Prod/Bundle Row # " & garyLRLine(I).TagID & " Bundle # " & J & " Of " & garyLRLine(I).LineBundleCount
                End If
                
                'now update the array position with the data inserted to replace the tag fields likee <THK> <BUNDLEID> etc...
                aryTagData(intTagCurrentRow) = strTagFile_AfterReplace
                Debug.Print aryTagData(intTagCurrentRow)
                
            Next 'J loop for 1 to bundle count for the item/row/product/bundle
        End If 'and of the If mixed product bundle check, where it's not mixed and is printing as part of the else =1 being false above
    Next 'I Loop for 0 to upper bound of teh garylrline array (lrline.pdb records)
    
    
    Debug.Print "udtLR Bundle Count=" & udtLR.BundleCount & "   vs   MixedCount+NonMixedCount=" & intTotalBundleCount_Mixed_NonMixed
    'now for each tag in teh aryTagData print the tag to the zebra/mobile/whatever printer.
    'FORCING "ALL" BUNDLES FOR NOW
    If intNumberOfTagsToPrint <= 0 Then intNumberOfTagsToPrint = udtLR.BundleCount
    
    For I = 0 To UBound(aryTagData)
        
        Dim fPrintThis As Boolean
        fPrintThis = False
        
        If intNumberOfTagsToPrint = 0 Or intNumberOfTagsToPrint = udtLR.BundleCount Then
            'print it.
            fPrintThis = True
        Else
            If I >= (udtLR.BundleCount - intNumberOfTagsToPrint) Then 'For example if 9 tags total, and want last three (ubound of arytagdata would be 8 so need position 6,7,8 ..subtracting # to print from the total bundles gives you 6..so >= this number..print, less...don't
                fPrintThis = True
            Else
                fPrintThis = False
            End If
        End If
        
        
        If fPrintThis = True Then
            Dim strCommSendToPrint As String
            strCommSendToPrint = aryTagData(I)
            Debug.Print "***TAG DATA ABOUT TO BE SENT TO COMM PORT *******" & vbCrLf & strCommSendToPrint
            
            If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3 PreSendCommData: " & vbCrLf & strCommSendToPrint
        
            lblReturnMessage.Caption = "Printing " & I & " OF " & udtLR.BundleCount
            lblReturnMessage.Refresh
    
            If IsNumeric(gSettings.RecPrintPauseCounterValue) = True Then
              If CLng(gSettings.RecPrintPauseCounterValue) >= 1 Then
                  'just continue
              Else
                  gSettings.RecPrintPauseCounterValue = "10"
              End If
                'do nothing
            Else
                gSettings.RecPrintPauseCounterValue = "10"
            End If
            
            For J = 1 To CLng(gSettings.RecPrintPauseCounterValue)
              'just waiting
                If fUseStatusLabel = True Then
                    lblStatus.Caption = "Print Pause " & J & " of " & gSettings.RecPrintPauseCounterValue
                    lblStatus.Refresh
                End If
            Next
            '****SEND TAG PRINT TO PRINTER?SERIAL PORT/BT COMM PORT
            If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3 PreSendCommData Immediately before bluetooth comm port send to Comm " & gSettings.BTComm
            If SC(strCommSendToPrint, "") = True Then
                'do nothing...nothing to send
            Else
                Debug.Print strCommSendToPrint
                BTSerial.Output = strCommSendToPrint
            End If
            
            If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3 PostSend Immediately AFTER bluetooth comm port send to Comm " & gSettings.BTComm
            lblReturnMessage.Caption = intTotalBundlesOnLoad & " Send to Printer...Completed!"
            lblReturnMessage.Refresh
        End If
    Next

    intError = 7000

    If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3 After Sends/End of Module"
    '''SocketWebData.Close
    BTSerial.PortOpen = False
    lblReturnMessage.Caption = "Printing Completed : Load " & udtLR.LoadAID
    lblReturnMessage.Refresh

intError = 8000

    PrintTag_AnyByReceiveID_V3 = True
    Exit Function
    
ErrorHandler:
On Error GoTo ErrorQuitAndExit
    MsgBox "Error (eLIT# " & intError & ") in PrintTag_AnyByReceiveID_V3 (CommPort=" & gSettings.BTComm & " LRDBInfo:  " & "LRID=" & udtLR.LRID & " Load# " & udtLR.LoadAID & " Print Method=" & strPrintType & vbCrLf & "Error# " & Err.Number & "-" & Err.Description
    PrintTag_AnyByReceiveID_V3 = False
    Exit Function
ErrorQuitAndExit:
    PrintTag_AnyByReceiveID_V3 = False
    Exit Function

End Function
Public Function PrintTag_ConnectToBTPrinter(strPrintType As String, ByRef BTSerial As AFSerial, ByRef strReturnErrorMessage As String) As Boolean
    Dim intError As Integer
    
On Error GoTo ErrorHandler
    If SC(strPrintType, "MOBILE") = True Or SC(strPrintType, "") = True Then 'Bluetooth/mobile printing
intError = 1
        BTSerial.InBufferSize = 2048
        BTSerial.PortOpen = False
        BTSerial.CommPort = gSettings.BTComm
    
intError = 2
        #If AppForge Then
            BTSerial.PortOpen = True
        #Else
            MsgBox "Skip Open Serial Port - Debug Mode"
        #End If
        
intError = 3
    Else
        MsgBox "Only mobile printing is currently enabled with V3 Tag printing / receiving"
        PrintTag_ConnectToBTPrinter = False
        Exit Function
    End If
    PrintTag_ConnectToBTPrinter = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error (eLIT# " & intError & ") in PrintTag_ConnectToBTPrinter (CommPort=" & gSettings.BTComm & " PrintType=" & strPrintType & vbCrLf & "Error# " & Err.Number & "-" & Err.Description
    PrintTag_ConnectToBTPrinter = False
    Exit Function

End Function

Private Function PrintTag_ReceiveV3_GetMixedTagArray(ByRef aryMixTags As aryStringType, ByRef fMixTagsFoundOnLoad As Boolean) As Boolean
    Dim I As Long
    
    
    Dim intProdLine As Integer, intBundleCount As Integer
On Error GoTo ErrorHandler

    
    ReDim aryMixTags.ary(RecV3MixTags.Last, 5)
    
    aryMixTags.ary(RecV3MixTags.MixGroupAID, 0) = "7"" XT"
    aryMixTags.ary(RecV3MixTags.MixGroupName, 0) = "7"" Cross Ties"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, 0) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, 0) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, 0) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupBundleCount, 0) = "0"
    
    
    aryMixTags.ary(RecV3MixTags.MixGroupAID, 1) = "6"" XT"
    aryMixTags.ary(RecV3MixTags.MixGroupName, 1) = "6"" Cross Ties"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, 1) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, 1) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, 1) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupBundleCount, 1) = "0"
    
    aryMixTags.ary(RecV3MixTags.MixGroupAID, 2) = "SW"
    aryMixTags.ary(RecV3MixTags.MixGroupName, 2) = "Switch Ties"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, 2) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, 2) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, 2) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupBundleCount, 2) = "0"
    
    aryMixTags.ary(RecV3MixTags.MixGroupAID, 3) = "BR"
    aryMixTags.ary(RecV3MixTags.MixGroupName, 3) = "Bridge Ties"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, 3) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, 3) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, 3) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupBundleCount, 3) = "0"
    
    aryMixTags.ary(RecV3MixTags.MixGroupAID, 4) = "CR"
    aryMixTags.ary(RecV3MixTags.MixGroupName, 4) = "Crossings"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, 4) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, 4) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, 4) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupBundleCount, 4) = "0"
    
    aryMixTags.ary(RecV3MixTags.MixGroupAID, 5) = "Other"
    aryMixTags.ary(RecV3MixTags.MixGroupName, 5) = "Other"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, 5) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, 5) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, 5) = "0"
    aryMixTags.ary(RecV3MixTags.MixGroupBundleCount, 5) = "0"
    
    fMixTagsFoundOnLoad = False
    
    For I = 0 To UBound(garyLRLine)
        
        Dim strGradeL2 As String
        Dim strThkAID As String, lngThkID As Long, strThickName As String
        
        Dim intMixGroupIndex As Integer
        strGradeL2 = ""
        
        strGradeL2 = Left(tcu(garyLRLine(I).Grade), 2)
        If garyLRLine(I).ThicknessID <= 0 Then
            Debug.Print "LRLineThick=" & garyLRLine(I).thickness & "  - LRLineThickID=" & garyLRLine(I).ThicknessID
            garyLRLine(I).ThicknessID = CLng(RR(GetThicknessData(garyLRLine(I).thickness, "HHAID", "ID", lngThkID, strThkAID, strThickName)))
            Debug.Print "**AFTER LOOKUP:   LRLineThick = " & garyLRLine(I).thickness & " - LRLineThickID = " & garyLRLine(I).ThicknessID
        End If
        
        strThkAID = GetThicknessData(tcu(garyLRLine(I).ThicknessID), "ID", "HHAID", lngThkID, strThkAID, strThickName)
        Debug.Print "Grade (Left2)=" & strGradeL2 & " - Thk=" & strThkAID
        
        If SC(strGradeL2, "XT") = True Then
            If SC(strThkAID, "7") = True Then
                intMixGroupIndex = 0
            ElseIf SC(strThkAID, "6") = True Then
                intMixGroupIndex = 1
            Else
                intMixGroupIndex = 5
            End If
            
        ElseIf SC(strGradeL2, "SW") = True Then
            intMixGroupIndex = 2
        ElseIf SC(strGradeL2, "BR") = True Then
            intMixGroupIndex = 3
        ElseIf SC(strGradeL2, "CR") = True Then
            intMixGroupIndex = 4
        Else
            intMixGroupIndex = 5
        End If
        
        aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, intMixGroupIndex) = CLng(aryMixTags.ary(RecV3MixTags.MixGroupPcsGrade, intMixGroupIndex)) + garyLRLine(I).LRLineXLong1
        aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, intMixGroupIndex) = CLng(aryMixTags.ary(RecV3MixTags.MixGroupPcsReject, intMixGroupIndex)) + garyLRLine(I).LRLineXLong2
        aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, intMixGroupIndex) = CLng(aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, intMixGroupIndex)) + garyLRLine(I).LRLineXLong1 + garyLRLine(I).LRLineXLong2
        aryMixTags.ary(RecV3MixTags.MixGroupBundleCount, intMixGroupIndex) = garyLRLine(I).LineBundleCount
        
        If CDbl(RR(aryMixTags.ary(RecV3MixTags.MixGroupPcsTotal, intMixGroupIndex))) > 0 Then
            fMixTagsFoundOnLoad = True
        End If
    Next
    
    PrintTag_ReceiveV3_GetMixedTagArray = True
    
    Exit Function

ErrorHandler:
    PrintTag_ReceiveV3_GetMixedTagArray = False
    MsgBox "Error in PrintTag_ReceiveV3_GetMixedTagArray " & Err.Number & "-" & Err.Description
    Exit Function

End Function
Public Function PrintTag_AnyByReceiveID_V3Mixed(udtLR As tLRRecord, ByRef aryMixedTags As aryStringType, ByRef intNextBundleID As Integer, _
                                            ByRef BTSerial As AFSerial, _
                                            ByRef strReturnMessage As String, _
                                            strTagFilePath As String, strTagFileName As String, intReturnMixedBundle_TotalCount As Integer, _
                                            Optional fUseLabel As Boolean, Optional lblStatus As AFLabel) As Boolean
    Dim I As Long, J As Long, K As Long
    Dim intError As Long
    Dim intTagCurrentRow As Integer, intGroupCurrentBundle As Integer
    
    Dim strLine As String
    
    Dim fFirstRow As Boolean
    Dim aryTagData() As String
    
    Dim fMixedProductsFoundOnLoad As Boolean
    Dim strReturnErrorMessage As String
   
    
On Error GoTo ErrorHandler
    
    intReturnMixedBundle_TotalCount = 0
    
    ReDim aryMixedTags.ary(RecV3MixTags.Last, 0)
    fMixedProductsFoundOnLoad = False
   
intError = 0
    'Connect to and open BTSerial Connection - Already connected form the V3 print module call that calls this one, so skip opening port
    
    'Get the mixed product tags array, if any on the load
    If PrintTag_ReceiveV3_GetMixedTagArray(aryMixedTags, fMixedProductsFoundOnLoad) = False Then
        'Error of some kind in module above, exit function
        PrintTag_AnyByReceiveID_V3Mixed = False
        strReturnErrorMessage = "Error occurred getting the mixed tag product array to print!"
        Exit Function
    End If
    
    If fMixedProductsFoundOnLoad = False Then 'no mixed tags to print, just exit the sub returning true
        PrintTag_AnyByReceiveID_V3Mixed = True
        Exit Function
    End If
    
    Dim intGroupBundleCount As Integer
    ReDim aryTagData(0)
    
    fFirstRow = False
    
    'Read Tag file Data to String for Printing Use
    strLine = fso_FileOpen_ReadOnly_ReturnFileTextOnly("", strTagFilePath, strTagFileName)
    Debug.Print "Tag Data From Tag File: " & strTagFilePath & " Name: " & strTagFileName & vbCrLf & strLine
    
    '*******LEAVING ROW ZERO  IN ARRAY BLANK TO SIMPLIFY CODE ******** WILL START PRINT AT ARRAY POSITION 1
    Dim aryProdGroup() As String
    Dim lngMixProdTotals As Long
    ReDim aryProdGroup(1, 0)
    
    ''*top section gets the prod group summary that prints on every mixed tag
    ReDim Preserve aryProdGroup(1, 0)  'used for printing summary on every ttag and to get mixed prod totals
    Dim fFirstProd As Boolean
    fFirstProd = False
    
    For I = 0 To UBound(aryMixedTags.ary, 2)
        If VC(aryMixedTags.ary(RecV3MixTags.MixGroupPcsTotal, I), 0) = True Then 'skip  it
            'don't include this one
        Else
            If fFirstProd = True Then
                ReDim Preserve aryProdGroup(1, UBound(aryProdGroup, 2) + 1) ' add new row
            Else
                fFirstProd = True 'use first zero position
            End If
            aryProdGroup(0, UBound(aryProdGroup, 2)) = aryMixedTags.ary(RecV3MixTags.MixGroupAID, I)
            aryProdGroup(1, UBound(aryProdGroup, 2)) = aryMixedTags.ary(RecV3MixTags.MixGroupPcsTotal, I)
            lngMixProdTotals = lngMixProdTotals + CLng(RR(aryMixedTags.ary(RecV3MixTags.MixGroupPcsTotal, I)))
        End If
    Next
    
    If UBound(aryProdGroup, 2) < 5 Then
        ReDim Preserve aryProdGroup(1, 5) 'just fill with blank rows which happens by the redim it'self to clear the ticket for prodgroup4/prodgroup5 etc if not used
    End If
        
    'Since there are mixed tags on load, need to go ahead and build the tag array, 1 for each mixed tag of each mixed type
    For I = 0 To UBound(aryMixedTags.ary, 2)
        intGroupBundleCount = 0
        intGroupBundleCount = CInt(RR(aryMixedTags.ary(RecV3MixTags.MixGroupBundleCount, I)))
        If intGroupBundleCount > 0 Then
            intTagCurrentRow = UBound(aryTagData)  'First row to insert this group into (Will Add+1 in loop below as we do it)
            ReDim Preserve aryTagData(UBound(aryTagData) + intGroupBundleCount)
        
            For J = 1 To intGroupBundleCount 'once for each bundle in this group
                intTagCurrentRow = intTagCurrentRow + 1
                
                intReturnMixedBundle_TotalCount = intReturnMixedBundle_TotalCount + 1
                
                aryTagData(intTagCurrentRow) = strLine 'initialize to the tag file data/file read results
                
                Dim strBundleIDLocal As String
                intNextBundleID = intNextBundleID + 1
                'Essentially bundle # is the Load# - the M0 (M Zero for Mixed bundle row, then a counter for next bundle in the mixed bundles)
                
                strBundleIDLocal = tcu(udtLR.LoadAID) & "-" & FixStringWidth("MX", 2, True, True) & FixStringWidth(CStr(intNextBundleID), 2, True, True)
                
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<BUNDLE>", strBundleIDLocal)
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<BUNDLEID", strBundleIDLocal)
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<12345678>", strBundleIDLocal)
                
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<LOAD>", tcu(udtLR.LoadAID))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<LOAD#>", tcu(udtLR.LoadAID))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<LOADAID>", tcu(udtLR.LoadAID))
                
                
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<LRDATE>", udtLR.LRXString1)
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<LOADDATE>", udtLR.LRXString1)
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<DATE>", udtLR.LRXString1)
                
                'HH User Logged In
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<PDINSPECTOR>", Trim(gstrUserName))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<PDINSP>", Trim(gstrUserName))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<INSP>", Trim(gstrUserName))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<USERNAME>", Trim(gstrUserName))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<USERID>", Trim(gstrUserName))
                
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<TOTALMIXEDPCS>", CStr(lngMixProdTotals))
                
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<PRODGROUP>", aryMixedTags.ary(RecV3MixTags.MixGroupName, I))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<PACKOFTOTALGROUP>", CStr(J) & " of " & CStr(intGroupBundleCount))
                aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<PACKOFTOTALALL>", CStr(intNextBundleID) & " of " & CStr(udtLR.BundleCount))
                
                'replace the summary on the ticket for <PRODGROUPX> variables and <PCSX> variables
                For K = 0 To UBound(aryProdGroup, 2)
                    aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<PRODGROUP" & CStr(K) & ">", aryProdGroup(0, K))
                    aryTagData(intTagCurrentRow) = Replace(aryTagData(intTagCurrentRow), "<PCS" & CStr(K) & ">", aryProdGroup(1, K))
                Next
                
                'now replace all the tags should be replaced as of now on this mixed product ticket
                Debug.Print "Tag Data(" & CStr(intTagCurrentRow) & ")- After Replaced Tag Fields: " & vbCrLf & aryTagData(intTagCurrentRow)
            Next
        End If
    Next
    
    
    
intError = 4
       
    For I = 1 To UBound(aryTagData)
        
        Dim strCommSendToPrint As String
        strCommSendToPrint = aryTagData(I)
        Debug.Print "***TAG DATA ABOUT TO BE SENT TO COMM PORT *******" & vbCrLf & strCommSendToPrint
            
        
        ''lblReturnMessage.Caption = "Print Mix Tag " & CStr(I) & " OF " & UBound(aryTagData)
         ''lblReturnMessage.Refresh
    
        If IsNumeric(gSettings.RecPrintPauseCounterValue) = True Then
          If CLng(gSettings.RecPrintPauseCounterValue) >= 1 Then
              'just continue
              
          Else
              gSettings.RecPrintPauseCounterValue = "10"
          End If
            'do nothing
        Else
            gSettings.RecPrintPauseCounterValue = "10"
        End If

        If SC(gSettings.DebugMode, "RECEIVEV3") = True Then
            MsgBox "Printing Mix Tag " & I & " of " & UBound(aryTagData) & " - RecPrintPauseCounterValue=" & CStr(gSettings.RecPrintPauseCounterValue)
        End If
        If fUseLabel = True Then
            If lblStatus.Visible = False Then lblStatus.Visible = True
            lblStatus.Caption = "Printing Mix Tag " & I & " Of ubound(aryTagData) PauseCount=" & CStr(gSettings.RecPrintPauseCounterValue)
            lblStatus.Refresh
        End If
        If SC(gSettings.DebugMode, "ReceiveV3") = True Then MsgBox "Print Pause Time set to " & CStr(CLng(gSettings.RecPrintPauseCounterValue))
        Dim lngPause As Long
        lngPause = 0
        '
        If SC(gSettings.DebugMode, "RECEIVEV3") = True Then MsgBox "Prepause start"
        
        For J = 1 To CLng(gSettings.RecPrintPauseCounterValue)
            If fUseLabel = True Then
                lblStatus.Caption = "Print Pause " & J & " of " & gSettings.RecPrintPauseCounterValue
                lblStatus.Refresh
            End If
          'just waiting
        Next
        
        If fUseLabel = True Then
            lblStatus.Caption = "Mixed Tag " & I & " of " & UBound(aryTagData) & " Sent to Print"
            lblStatus.Refresh
        End If
        
        
        If SC(gSettings.DebugMode, "RECEIVEV3") = True Then MsgBox "PrePause End"
        
        '****SEND TAG PRINT TO PRINTER?SERIAL PORT/BT COMM PORT
        If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3Mixed PreSendCommData Immediately before bluetooth comm port send to Comm " & gSettings.BTComm
        If SC(strCommSendToPrint, "") = True Then
            'do nothing...nothing to send
        Else
            BTSerial.Output = strCommSendToPrint
        End If
        
        If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3Mixed PostSend Immediately AFTER bluetooth comm port send to Comm " & gSettings.BTComm
        ''lblReturnMessage.Caption = "Mixed " & CStr(I) & " of " & UBound(aryTagData) & " Send to Printer...Completed!"
        ''lblReturnMessage.Refresh
    
    Next

    intError = 7000

    If SC(gSettings.DebugMode, "P4T") = True Then MsgBox "DebugMode: P4T  PrintTag_AnyByReceiveID_V3Mixed After Sends/End of Module"
    
'''    Do not close port because it's being used in the partent module for non-mixed tags also
    '''SocketWebData.Close
    '''BTSerial.PortOpen = False
    
    ''lblReturnMessage.Caption = "Mix Tag Printing Completed : Load " & udtLR.LoadAID
    ''lblReturnMessage.Refresh

intError = 8000
    'intnextbundleid will be returned to show the X of 10 count of the following bundles as well.
    
    PrintTag_AnyByReceiveID_V3Mixed = True
    Exit Function
    
ErrorHandler:
On Error GoTo ErrorQuitAndExit
    MsgBox "Error (eLIT# " & intError & ") in PrintTag_AnyByReceiveID_V3Mixed (CommPort=" & gSettings.BTComm & " LRDBInfo:  " & "LRID=" & udtLR.LRID & " Load# " & udtLR.LoadAID & " Print Method=" & "MOBILE" & vbCrLf & "Error# " & Err.Number & "-" & Err.Description
    PrintTag_AnyByReceiveID_V3Mixed = False
    Exit Function
ErrorQuitAndExit:
    PrintTag_AnyByReceiveID_V3Mixed = False
    Exit Function

End Function

Public Function PrintTag_AnyByReceiveID_ReplaceTags_LRPrint(ByRef udtLR As tLRRecord, udtLRLine As tLRLineRecord, _
                                                            ByVal strLineDataIn As String, ByRef strDataOut As String, _
                                                            intRow_BundleIndex As Integer, intRow_BundleCountOfTotal As Integer) As Boolean
    Dim I As Long
    Dim strLoadAID As String
On Error GoTo ErrorHandler

    strDataOut = strLineDataIn
    
    'Start replacing tag fields in file with actual recrod values
    '**BUNDLEID **** V3
    'Bundle ID will be Load#-ProdRow#-# of Bundle of that prod (two digit for last two parts ..both)
    'Load R13356 with 3 products on it, first on has 2 bundles, 2nd has 8...and so on...Bundle1  R13356-0101,R13356-0102 (first two bundles of product 1), R13356-0201,R13356-0202,R13356-0203,...R13356-0208,  NS AO ON
    Dim strBundleIDLocal As String
    
    strBundleIDLocal = tcu(udtLR.LoadAID) & "-" & FixStringWidth(udtLRLine.TagID, 2, True, True) & FixStringWidth(CStr(intRow_BundleIndex), 2, True, True)

    strDataOut = Replace(strDataOut, "<BUNDLE>", strBundleIDLocal)
    strDataOut = Replace(strDataOut, "<BUNDLEID", strBundleIDLocal)
    strDataOut = Replace(strDataOut, "<12345678>", strBundleIDLocal)
    
    strLoadAID = udtLR.LoadAID
    If SC(udtLRLine.LRLineX1, "1") = True Then 'mixed product tag
        'stick in load number for now
        udtLR.LoadAID = udtLR.LoadAID & "-Mix"
    End If
    
    strDataOut = Replace(strDataOut, "<LOAD>", tcu(udtLR.LoadAID))
    strDataOut = Replace(strDataOut, "<LOAD#>", tcu(udtLR.LoadAID))
    strDataOut = Replace(strDataOut, "<LOADAID>", tcu(udtLR.LoadAID))
    udtLR.LoadAID = strLoadAID
    'sinc this is passed by ref, now set it back to the correct if mixed product
    strDataOut = Replace(strDataOut, "<PRINTDATETIME>", FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbShortTime))
    strDataOut = Replace(strDataOut, "<TIME>", FormatDateTime(Now, vbShortTime))
    
    strDataOut = Replace(strDataOut, "<BOL>", tcu(udtLR.BOLID))
    strDataOut = Replace(strDataOut, "<BOLID>", tcu(udtLR.BOLID))
    strDataOut = Replace(strDataOut, "<REF#>", tcu(udtLR.BOLID))
            
    strDataOut = Replace(strDataOut, "<VENDOR>", udtLR.VendorName)
    strDataOut = Replace(strDataOut, "<VENDORNAME>", udtLR.VendorName)
    strDataOut = Replace(strDataOut, "<VENDORAID>", GetOrgData(CStr(udtLR.VendorID), "ID", "VENDOR", "HHAID"))
            
            
    strDataOut = Replace(strDataOut, "<PONUMBER>", tcu(udtLR.PONumber))
    
    strDataOut = Replace(strDataOut, "<LOCATION>", tcu(udtLR.Location))
    strDataOut = Replace(strDataOut, "<LOCAID>", tcu(udtLR.Location))
    strDataOut = Replace(strDataOut, "<LOC>", tcu(udtLR.Location))
    
    strDataOut = Replace(strDataOut, "<LOCNAME>", (GetLocData(0, udtLR.Location, "LOCNAME")))
    strDataOut = Replace(strDataOut, "<LOCDESC>", (GetLocData(0, udtLR.Location, "LOCNAME")))
    
    
    strDataOut = Replace(strDataOut, "<DATE>", udtLR.LRXString1)
    strDataOut = Replace(strDataOut, "<LRDATE>", udtLR.LRXString1)
    strDataOut = Replace(strDataOut, "<LOADDATE>", udtLR.LRXString1)
    
    strDataOut = Replace(strDataOut, "<BUNDLECOUNT>", FixStringWidth(CStr(udtLR.BundleCount), 3, True, False))
    
    strDataOut = Replace(strDataOut, "<THKXWIDTHXLEN>", tcu(udtLRLine.thickness) & " X " & tcu(CStr(udtLRLine.AvgWidthBoard)) & " X " & tcu(CStr(udtLRLine.AvgLenBoard)))
    strDataOut = Replace(strDataOut, "<THKWIDTHLEN>", tcu(udtLRLine.thickness) & " X " & tcu(CStr(udtLRLine.AvgWidthBoard)) & " X " & tcu(CStr(udtLRLine.AvgLenBoard)))
    
    strDataOut = Replace(strDataOut, "<THK>", GetThicknessData(CStr(udtLRLine.ThicknessID), "ID", "THICKNESS"))
    strDataOut = Replace(strDataOut, "<THICKNESS>", GetThicknessData(CStr(udtLRLine.ThicknessID), "ID", "THICKNESS"))
    strDataOut = Replace(strDataOut, "<THICKNESSNAME>", GetThicknessData(CStr(udtLRLine.ThicknessID), "ID", "THICKNESS"))
    
    strDataOut = Replace(strDataOut, "<THICKNESSAID>", GetThicknessData(CStr(udtLRLine.ThicknessID), "ID", "HHAID"))
    strDataOut = Replace(strDataOut, "<THKAID>", GetThicknessData(CStr(udtLRLine.ThicknessID), "ID", "HHAID"))
    strDataOut = Replace(strDataOut, "<THICKAID>", GetThicknessData(CStr(udtLRLine.ThicknessID), "ID", "HHAID"))
    
    
    strDataOut = Replace(strDataOut, "<SPECIESAID>", GetSpeciesData(CStr(udtLRLine.SpeciesID), "ID", "HHAID"))
    strDataOut = Replace(strDataOut, "<SPECIESNAME>", GetSpeciesData(CStr(udtLRLine.SpeciesID), "ID", "SPECIESNAME")) ' these two shoudl be same, allows name/desc/etc for ease of use/memory
    strDataOut = Replace(strDataOut, "<SPECIESDESC>", GetSpeciesData(CStr(udtLRLine.SpeciesID), "ID", "SPECIESDESC")) 'these two should be same
    
    strDataOut = Replace(strDataOut, "<GRADEAID>", GetGradeData(CStr(udtLRLine.GradeID), "ID", "HHAID"))
    strDataOut = Replace(strDataOut, "<GRADENAME>", GetGradeData(CStr(udtLRLine.GradeID), "ID", "GRADE"))
    strDataOut = Replace(strDataOut, "<GRADEDESC>", GetGradeData(CStr(udtLRLine.GradeID), "ID", "GRADE"))
    
    strDataOut = Replace(strDataOut, "<GRADE>", GetGradeData(CStr(udtLRLine.GradeID), "ID", "GRADE"))
    
    strDataOut = Replace(strDataOut, "<WIDTH>", CStr(udtLRLine.AvgWidthBoard))
    strDataOut = Replace(strDataOut, "<AWIDTH>", CStr(udtLRLine.AvgWidthBoard))
    strDataOut = Replace(strDataOut, "<WIDTHAID>", CStr(udtLRLine.AvgWidthBoard))
        
    strDataOut = Replace(strDataOut, "<LEN>", CStr(udtLRLine.AvgLenBoard))
    strDataOut = Replace(strDataOut, "<LENGTH>", CStr(udtLRLine.AvgLenBoard))
    
        

    strDataOut = Replace(strDataOut, "<#OF#INPROD>", CStr(intRow_BundleIndex) & " of " & CStr(udtLRLine.LineBundleCount))
    strDataOut = Replace(strDataOut, "<PACK>", CStr(intRow_BundleCountOfTotal) & " of " & CStr(udtLR.BundleCount))
    strDataOut = Replace(strDataOut, "<PACK#>", CStr(intRow_BundleCountOfTotal) & " of " & CStr(udtLR.BundleCount))
    

    strDataOut = Replace(strDataOut, "<PACKFOOTAGE>", CStr(udtLRLine.Footage))
    strDataOut = Replace(strDataOut, "<FTG>", CStr(udtLRLine.Footage))
    strDataOut = Replace(strDataOut, "<FOOTAGE>", CStr(udtLRLine.Footage))
    strDataOut = Replace(strDataOut, "<BDFT>", CStr(udtLRLine.Footage))
            
           
    'HH User Logged In
    strDataOut = Replace(strDataOut, "<PDINSPECTOR>", Trim(gstrUserName))
    strDataOut = Replace(strDataOut, "<PDINSP>", Trim(gstrUserName))
    strDataOut = Replace(strDataOut, "<INSP>", Trim(gstrUserName))
    strDataOut = Replace(strDataOut, "<USERNAME>", Trim(gstrUserName))
    strDataOut = Replace(strDataOut, "<USERID>", Trim(gstrUserName))
           
              
    '''strDataOut = Replace(strDataOut, "<COLORAID", gudtPD.ColorAID)
    '**PCS/REJECTS - On Grade Pieces, Reject Pieces, Total Pieces
    strDataOut = Replace(strDataOut, "<PCS>", CStr(udtLRLine.LRLineXLong1)) ' PCS
    strDataOut = Replace(strDataOut, "<PIECES>", CStr(udtLRLine.LRLineXLong1)) ' PCS
            
    strDataOut = Replace(strDataOut, "<REJECTS>", CStr(udtLRLine.LRLineXLong2)) ' PCS REJECTS
    strDataOut = Replace(strDataOut, "<REJECTPIECES>", CStr(udtLRLine.LRLineXLong2)) ' PCS REJECTS
    strDataOut = Replace(strDataOut, "<PIECESREJECT>", CStr(udtLRLine.LRLineXLong2)) ' PCS REJECTS
            
    'Total Pieces for Product
    strDataOut = Replace(strDataOut, "<TOTALPCS>", CStr(udtLRLine.LRLineXLong2 + udtLRLine.LRLineXLong1)) ' PCS REJECTS + PCS ON GRADE
    strDataOut = Replace(strDataOut, "<TOTALPIECES>", CStr(udtLRLine.LRLineXLong2 + udtLRLine.LRLineXLong1)) ' PCS REJECTS + PCS ON GRADE
    strDataOut = Replace(strDataOut, "<PCSTOTAL>", CStr(udtLRLine.LRLineXLong2 + udtLRLine.LRLineXLong1)) ' PCS REJECTS + PCS ON GRADE
    strDataOut = Replace(strDataOut, "<PIECESTOTAL>", CStr(udtLRLine.LRLineXLong2 + udtLRLine.LRLineXLong1)) ' PCS REJECTS + PCS ON GRADE
    strDataOut = Replace(strDataOut, "<LAYERS>", CStr(udtLRLine.LRLineXLong2 + udtLRLine.LRLineXLong1)) ' PCS REJECTS + PCS ON GRADE
    
    If SC(udtLRLine.LRLineX1, "1") = True Then
        
        If (InStr(strDataOut, "<MIXEDPRODUCTS>") > 0 Or InStr(strDataOut, "<MIXPRODUCTS>") > 0 Or InStr(strDataOut, "<MIXED>") > 0) Then
        'mixed product bundles, print all descriptions
        
    Dim strMix As String
            strMix = ""
            'First put this row in, then the others set as mixed
            strMix = udtLRLine.Species & " " & udtLRLine.Grade & " " & tcu(udtLRLine.thickness) & " X " & tcu(CStr(udtLRLine.AvgWidthBoard)) & " X " & tcu(CStr(udtLRLine.AvgLenBoard)) & " Pcs: " & CStr(udtLRLine.LRLineXLong1 + udtLRLine.LRLineXLong2)
        
            For I = 0 To UBound(garyLRLine)
                If SC(garyLRLine(I).LRLineX1, "1") = True And garyLRLine(I).LRLineID <> udtLRLine.LRLineID Then
                    'Add this one too
                    strMix = strMix & vbCrLf & garyLRLine(I).Species & " " & garyLRLine(I).Grade & " " & tcu(garyLRLine(I).thickness) & " X " & tcu(CStr(garyLRLine(I).AvgWidthBoard)) & " X " & tcu(CStr(garyLRLine(I).AvgLenBoard)) & " Pcs: " & CStr(garyLRLine(I).LRLineXLong1 + garyLRLine(I).LRLineXLong2)
                End If
            Next
            
            strDataOut = Replace(strDataOut, "<MIXEDPRODUCTS>", strMix)
            strDataOut = Replace(strDataOut, "<MIXPRODUCTS>", strMix)
            strDataOut = Replace(strDataOut, "<MIXED>", strMix)
        End If
    End If
    
    PrintTag_AnyByReceiveID_ReplaceTags_LRPrint = True
    Debug.Print "Tag to Print prior to adding to aryTagData() *****************" & vbCrLf & strDataOut & vbCrLf & vbCrLf & "****** END OF TAG DATA TO PRINT **************"
    Exit Function
    
ErrorHandler:
On Error GoTo ErrorQuitAndExit
    PrintTag_AnyByReceiveID_ReplaceTags_LRPrint = False
    
    MsgBox "Error in PrintTag_AnyByReceiveID_ReplaceTags_LRPrint " & vbCrLf & "Error# " & Err.Number & "-" & Err.Description
    PrintTag_AnyByReceiveID_ReplaceTags_LRPrint = False
    Exit Function
ErrorQuitAndExit:
    PrintTag_AnyByReceiveID_ReplaceTags_LRPrint = False
    Exit Function

End Function
'Public Function dbLengthBulkEditDeveloperOnly()
'
'    Dim udt As tLengthRecord
'    Dim dbTemp As Long
'    Dim I As Integer
'    Dim intError As Integer
'    Dim FSO As CFileManager
'On Error GoTo ErrorHandler
'    Exit Function
'    Set FSO = New CFileManager
'
'    OpenLengthDatabase
'    PDBSetSortFields dbLength, tLengthDatabaseFields.LengthID_Field
'    PDBMoveFirst dbLength
'    ReDim gudtLenNew(0)
'    If PDBNumRecords(dbLength) = 0 Then
'        strCommSendBox "ERROR EXIT SUB"
'        Exit Function
'    End If
'
'    ReDim gudtLenNew(PDBNumRecords(dbLength) - 1)
'
'    PDBBulkRead dbLength, PDBNumRecords(dbLength), VarPtr(gudtLenNew(0))
'    CloseLengthDatabase
'    dbLength = 0
'
'
' '   FSO.DeleteFile (gstrPDBPath & "\length.pdb")
'  '          CreateDatabasePDB dbLength, "length", Length_Schema
'   '         PDBClose dbLength
'
'    OpenLengthDatabase
'    For I = 0 To UBound(gudtLenNew)
'        If SC(gudtLenNew(I).LengthUnit, "FEET") = True Then
'            gudtLenNew(I).LengthUnit = "FT"
'        ElseIf SC(gudtLenNew(I).LengthUnit, "Inches") = True Then
'            gudtLenNew(I).LengthUnit = "IN"
'        ElseIf SC(gudtLenNew(I).LengthUnit, "") = True Then
'            gudtLenNew(I).LengthUnit = "FT"
'        End If
'
'        If InStr(gudtLenNew(I).HHAID, "'") > 0 Then
'            gudtLenNew(I).HHAID = "" & tcu(Replace(gudtLenNew(I).HHAID, "'", ""))
'        ElseIf InStr(gudtLenNew(I).HHAID, """") > 0 Then
'            gudtLenNew(I).HHAID = "0" & tcu(Replace(gudtLenNew(I).HHAID, """", ""))
'        End If
'        PDBMoveFirst dbLength
'        PDBFindRecordByField dbLength, tLengthDatabaseFields.LengthID_Field, gudtLenNew(I).LengthID
'        If PDBGetLastError(dbLength) = 0 Then
'            PDBEditRecord dbLength
'            Call WriteLengthRecord(gudtLenNew(I))
'            PDBUpdateRecord dbLength
'        End If
'
'    Next
'
'    CloseLengthDatabase
'    dbLength = 0
'
'    Exit Function
'ErrorHandler:
'    strCommSendBox "ERROR'"
'
'
'End Function

'****************
' Main Function *
'****************
Function SpellNumber(ByVal MyNumber As String) As String
    Dim Dollars As String, Cents As String, temp As String
    Dim DecimalPlace As Integer, COUNT As Integer
    Dim aryPlace() As String
    
On Error GoTo ErrorHandler
    If SC(MyNumber, "") = True Then
        SpellNumber = ""
        Exit Function
    End If
    
    ReDim aryPlace(9) As String
    aryPlace(2) = " Thousand "
    aryPlace(3) = " Million "
    aryPlace(4) = " Billion "
    aryPlace(5) = " Trillion "
    
    ' String representation of amount
    MyNumber = Trim((MyNumber))
    If SC(MyNumber, "") = True Then Exit Function
    
    'Replace all commas, dollar signs, etc.
    MyNumber = CStr(RR(MyNumber))
    
    ' Position of decimal place 0 if none
    DecimalPlace = InStr(MyNumber, ".")
    'Convert cents and set MyNumber to dollar amount
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
 
    COUNT = 1
    Do While MyNumber <> ""
       temp = GetHundreds(Right(MyNumber, 3))
       If temp <> "" Then Dollars = temp & aryPlace(COUNT) & Dollars
          If Len(MyNumber) > 3 Then
             MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        COUNT = COUNT + 1
    Loop
 
    Select Case Dollars
        Case ""
            Dollars = "No Dollars"
        Case "One"
            Dollars = "One Dollar"
        Case Else
            Dollars = Dollars & " Dollars"
    End Select
 
    Select Case Cents
        Case ""
            Cents = " and No Cents"
        Case "One"
            Cents = " and One Cent"
        Case Else
            Cents = " and " & Cents & " Cents"
    End Select
 
    SpellNumber = Dollars & Cents
Exit Function
ErrorHandler:
    MsgBox "Error Converting " & MyNumber & " To Word Amounts " & Err.Number & "-" & Err.Description
    SpellNumber = "ERROR-" & MyNumber
    Exit Function
End Function
 
'*******************************************
' Converts a number from 100-999 into text *
'*******************************************
Function GetHundreds(ByVal MyNumber As String) As String
    Dim Result As String
 
    If IsNumeric(MyNumber) = True Then
        If CDbl(MyNumber) = 0 Then Exit Function
    Else
        Exit Function
    End If
    
    MyNumber = Right("000" & MyNumber, 3)
 
    'Convert the hundreds place
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
 
    'Convert the tens and ones place
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
 
    GetHundreds = Result
End Function
 
'*********************************************
' Converts a number from 10 to 99 into text. *
'*********************************************
Function GetTens(TensText As String) As String
    Dim Result As String
 
    Result = ""           'null out the temporary function value
    If IsNumeric(TensText) = True Then
        'just continue
    Else
        GetTens = "Zero"
        Exit Function
    End If
    
    If CDbl(Left(TensText, 1)) = 1 Then   ' If value between 10-19
        Select Case CInt(TensText)
            Case 10: Result = "Ten"
            Case 11: Result = "Eleven"
            Case 12: Result = "Twelve"
            Case 13: Result = "Thirteen"
            Case 14: Result = "Fourteen"
            Case 15: Result = "Fifteen"
            Case 16: Result = "Sixteen"
            Case 17: Result = "Seventeen"
            Case 18: Result = "Eighteen"
            Case 19: Result = "Nineteen"
            Case Else
        End Select
      Else                                 ' If value between 20-99
        Select Case CInt(Left(TensText, 1))
            Case 2: Result = "Twenty "
            Case 3: Result = "Thirty "
            Case 4: Result = "Forty "
            Case 5: Result = "Fifty "
            Case 6: Result = "Sixty "
            Case 7: Result = "Seventy "
            Case 8: Result = "Eighty "
            Case 9: Result = "Ninety "
            Case Else
        End Select
         Result = Result & GetDigit _
            (Right(TensText, 1))  'Retrieve ones place
      End If
      GetTens = Result
   End Function
 
'*******************************************
' Converts a number from 1 to 9 into text. *
'*******************************************
Function GetDigit(Digit As String) As String
    If IsNumeric(Digit) = True Then
        'just continue
    Else
        'just exit I guess
        Exit Function
    End If
    
    Select Case CInt(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

Public Function GetAvgLengthMinMax(strAvgLength As String, strMinMax As String) As Integer
    Dim I As Integer
    Dim fFound As Boolean
On Error GoTo ErrorHandler

    
    '''Created by BJP 10/13/2016 - gets the min and max length to pass into GetLengthRangeData module
    
    If IsNumeric(strAvgLength) = True And SC(strAvgLength, "") = False Then
        If InStr(strAvgLength, ".") > 0 Then
            If SC(strMinMax, "MAX") = True Then
                GetAvgLengthMinMax = CInt(Round(CDbl(strAvgLength) + 0.5, 0))
            ElseIf SC(strMinMax, "MIN") = True Then
                GetAvgLengthMinMax = CInt(Round(CDbl(strAvgLength) - 0.49, 0))
            End If
        Else
            'This is a whole number
            If SC(strMinMax, "MAX") = True Then
                GetAvgLengthMinMax = CInt(Round(CDbl(strAvgLength), 0))
            ElseIf SC(strMinMax, "MIN") = True Then
                GetAvgLengthMinMax = CInt(Round(CDbl(strAvgLength), 0))
            End If
        End If
    Else
        ''Blank or Non-Numeric - Shouldn't happen.
    End If

    
    Exit Function
       
ErrorHandler:
    MsgBox "Error in GetAvgLengthMinMax " & Err.Number & "-" & Err.Description
    Exit Function
       
End Function

Public Function ShowLookupForm(strType As String, frmReturn As Form, txtReturn As AFTextBox, Optional strFilter As String) As Boolean

On Error GoTo ErrorHandler
    frmLookup.pstrType = strType
    frmLookup.pstrFilter = strFilter
    Set frmLookup.pfrmReturn = frmReturn
    Set frmLookup.txtReturn = txtReturn
    frmLookup.Show

    ShowLookupForm = True
    Exit Function
ErrorHandler:
    ShowLookupForm = False
    MsgBox "Error in ShowLookupForm " & Err.Number & "-" & Err.Description
    Exit Function
    
End Function

Public Function ShowListForm(strType As String, Optional strAIDToSelect As String) As Boolean

On Error GoTo ErrorHandler


        frmList.pstrType = strType
        If SC(strAIDToSelect, "") = False Then
            frmList.pstrAIDToSelect = tcu(strAIDToSelect)
        Else
            frmList.pstrAIDToSelect = ""
        End If
        
        frmList.Show
        
    
Exit Function
ErrorHandler:
    ShowListForm = False
    MsgBox "Error in ShowListForm Type=" & strType & " " & Err.Number & "-" & Err.Description
    Exit Function
    
    
End Function


Public Function CheckForHandheldLicense(aryRomSerial As aryStringType) As Boolean

On Error GoTo ErrorHandler
    
    ReDim aryRomSerial.ary(2, 600)
        
    
    
        aryRomSerial.ary(0, 0) = "74233DF43F10076D1"
        aryRomSerial.ary(0, 1) = "D4A110744D10076D1"
        aryRomSerial.ary(0, 3) = "7E5151F44410076D1"
        aryRomSerial.ary(0, 4) = "74E261E44810076D1"
        aryRomSerial.ary(0, 5) = "01600D740310076D1" 'ITTW01
        aryRomSerial.ary(0, 6) = "0BF008842110076D1" 'DEMO 7900 UNIT
        aryRomSerial.ary(0, 7) = "D4C110744910076D1" 'HH1018
        aryRomSerial.ary(0, 8) = "D6D361D44110076D1" 'HH1019
        aryRomSerial.ary(0, 9) = "D6D361D44110076D1" 'HH1019
        aryRomSerial.ary(0, 10) = "0AB001741810076D1" 'trucS9
        aryRomSerial.ary(0, 11) = "03100FF42110076D1" 'HH1007
        aryRomSerial.ary(0, 12) = "0AB001741810076D1" 'GPTRUCS9'
        aryRomSerial.ary(0, 13) = "0D1000143710076D1" 'GPTRUCS6
        aryRomSerial.ary(0, 14) = "D4F110744E10076D1" 'HH assigned to sherwood
        aryRomSerial.ary(0, 15) = "4EB482808210076D1" 'HH1001
        aryRomSerial.ary(0, 16) = "0C6001441210076D1" 'HH1002
        aryRomSerial.ary(0, 17) = "0DB001741710076D1" 'HH1003
        aryRomSerial.ary(0, 18) = "0B8000343B10076D1" 'HH1004
        aryRomSerial.ary(0, 19) = "0C1000143310076D1"
        aryRomSerial.ary(1, 19) = "9500"
        
        aryRomSerial.ary(0, 20) = "0D8000343410076D1" 'HH1010
        aryRomSerial.ary(0, 21) = "042002A41410076D1" 'HHP Owned DEVICE
        aryRomSerial.ary(0, 22) = "09D00FB42C10076D1" 'HH009
        aryRomSerial.ary(0, 23) = "038000343010076D1" 'HH1011
        aryRomSerial.ary(0, 24) = "C3F663644310076D1" 'New One 0192412
        aryRomSerial.ary(0, 25) = "C3B663544710076D1" '
        aryRomSerial.ary(0, 26) = "7D6523344810076D1"
        aryRomSerial.ary(0, 27) = "C34663544510076D1" '
        aryRomSerial.ary(0, 28) = "0B8000343A10076D1" ' Handled HH. delete this serial once HH issold!
        aryRomSerial.ary(0, 29) = "C31663544A10076D1" '
        aryRomSerial.ary(0, 30) = "00600D840E10076D1" '
        aryRomSerial.ary(0, 31) = "C31663544A10076D1" '
        aryRomSerial.ary(0, 32) = "D8622AD44F10076D1" '
        aryRomSerial.ary(0, 33) = "7DA523344E10076D1" '
        
        aryRomSerial.ary(0, 34) = "440037003600300030000000-4137303000" ' HHP 7600 ITTW02
        aryRomSerial.ary(1, 34) = "7600"
        
        aryRomSerial.ary(0, 35) = "7D0523344810076D1" '
        aryRomSerial.ary(0, 36) = "0DC004D41110076D1" '
        aryRomSerial.ary(0, 37) = "7D7523344310076D1" 'HH1023
        aryRomSerial.ary(0, 60) = "0DB00E440710076D1" 'SN 0053270 Baillie Titusville, Updated 9/2/2010
        aryRomSerial.ary(0, 61) = "40A663344E10076D1" 'SN 213892 - Baillie Titusville, Updated 9/2/2010
        aryRomSerial.ary(0, 62) = "19A86EE45710076D1"
        aryRomSerial.ary(0, 63) = "8C1061445110076D1"
        aryRomSerial.ary(0, 64) = "910151546310076D1"
        aryRomSerial.ary(0, 65) = "91C151646310076D1"
        aryRomSerial.ary(0, 66) = "8C4161445510076D1"
        aryRomSerial.ary(0, 67) = "88D651A46410076D1"
        aryRomSerial.ary(0, 68) = "889651A46210076D1"     ' Milton Timber
        aryRomSerial.ary(0, 69) = "888651A46A10076D1"    ' Leer
        aryRomSerial.ary(0, 70) = "882651B46010076D1"
        aryRomSerial.ary(0, 71) = "3FD62A245710076D1"
        aryRomSerial.ary(0, 72) = "889651A46010076D1"              ' 47 Lumber
        aryRomSerial.ary(0, 73) = "3FF62A245F10076D1" '56KEY 9500LOP-432C30E  SERIAL: 0286123
        aryRomSerial.ary(0, 74) = "889651B46610076D1" '35KEY 9500L0P-422C30E  SERIAL: 0286088
        aryRomSerial.ary(0, 75) = "886651A46D10076D1" '43key 9500L0P-412C30E  SERIAL: 0286145
        aryRomSerial.ary(0, 76) = "886651B46910076D1" '35key 9500L0P-422C30E  SERIAL: 0286073
        aryRomSerial.ary(0, 77) = "886651B46710076D1" '35key 9500L0P-422C30E  SERIAL: 0286072
        aryRomSerial.ary(0, 78) = "88C651A46D10076D1" '43key 9500L0P-412C30E  SERIAL: 0286156
        aryRomSerial.ary(0, 79) = "88E651A46B10076D1" '35KEY 9500L0P-422C30E  SERIAL: 0286089
        
        aryRomSerial.ary(0, 80) = "1002A4016245D251A800-0050BF7A60E2" 'HH1304 'Atlanta Hardwoods Bruce manalan Symbol Device #1
        aryRomSerial.ary(1, 80) = "SYMBOL"
        
        aryRomSerial.ary(0, 81) = "077192003046B9B1D800-0050BF7A60E2" 'HH1305 Atlanta Hardwoods Part #  MC9090-GJ0HBFGA2WR   S/N 8045000502526 3/18/08
        aryRomSerial.ary(1, 81) = "SYMBOL"
        
        aryRomSerial.ary(0, 82) = "0392A9010347CB911800-0050BF7A60E2" 'HH1306 Atlanta Hardwoods Part # MC9090-GJ0HBFGA2WR   s/n 8010000505870   3/18/08
        aryRomSerial.ary(1, 82) = "SYMBOL"
        
        aryRomSerial.ary(0, 83) = "6B0353746910076D1" '35KEY 9500L0P-422C30E Serial: 0306570  Arrived 3/18/08
        aryRomSerial.ary(0, 84) = "3FA62A245A10076D1" '56KEY 9500L0P-432C30E  Serial: 0286103 Arrived 3/18/08
        aryRomSerial.ary(0, 85) = "6BC353746910076D1" '43KEY 9500L0P-412C30E  Serial: 0306767 Arrived 3/18/08
        aryRomSerial.ary(0, 86) = "6B6353846C10076D1" '35Key 9500L0P-422C30E  Serial: 0307458  Roan Mtn, TN
        aryRomSerial.ary(0, 87) = "6BF353946010076D1" '35Key 9500L0P-422C30E  Serial:  0307464 'Currently MNoland's Unit
        aryRomSerial.ary(0, 88) = "6b8353946710076d1" '9500L0P-422c30e (35 key)  S/n 0307465"
        aryRomSerial.ary(0, 89) = "6bd353946010076d1" '9500L0P-422c30e (35 key) s/n 0307457"
        aryRomSerial.ary(0, 90) = "6b0353746f10076d1" '9500L0P-422c30e (35 key) s/n 0306760"
        aryRomSerial.ary(0, 91) = "6bc353746d10076d1" '9500L0P-422c30e (35 key) s/n 0307456"
        aryRomSerial.ary(0, 92) = "6b4353746910076d1" '9500L0P-422c30e (35 key) s/n 0306574"
        aryRomSerial.ary(0, 93) = "6b2353746010076d1" '9500L0P-422c30e (35 key) s/n 0307462"
        aryRomSerial.ary(0, 94) = "6be353946a10076d1" '9500L0P-422c30e (35 key) s/n 0307461"
        aryRomSerial.ary(0, 95) = "6bc353746e10076d1" '9500L0P-422c30e (35 key) s/n 0306582"
        aryRomSerial.ary(0, 96) = "6b2353746a10076d1" '9500L0P-422c30e (35 key) s/n 0306572"
        aryRomSerial.ary(0, 97) = "6bd353946810076d1" '9500L0P-422c30e (35 key) s/n 0307459"
        aryRomSerial.ary(0, 98) = "6bf353746810076d1" '9500L0P-422c30e (35 key) s/n 0306575"
        aryRomSerial.ary(0, 99) = "3f662a245010076d1" '9500L0P-432C30E (56 key) s/n 0286115"
        aryRomSerial.ary(0, 100) = "DUPLICATED" '9500L0P-432C30E (56 key) s/n 0286105"  HH 1084 'Coastal Spartansburg Receiving.  Model#: 9500L0P-432C30E  Serial#: 286105  Manu: 10/26/2007  Install 9/23/2008  HH1084
        aryRomSerial.ary(0, 101) = "3f262a245c10076d1" '9500L0P-432C30E (56 key) s/n 0286122" 'Coastal Spartansburg Receiving.  Model#: 9500L0P-432C30E  Serial#: 286122  Manu: 10/26/2007  Install 9/23/2008  HH1078
        aryRomSerial.ary(0, 102) = "3f162a245210076d1" '9500L0P-432C30E (56 key) s/n 0286117"
        aryRomSerial.ary(0, 103) = "3f062a245a10076d1" '9500L0P-432C30E (56 key) s/n 0286120"
        
        aryRomSerial.ary(0, 104) = "6bb353746110076d1" '9500L0P-412C30E (43 key) s/n 0306547" Pan Pac 9500 56 Key
        aryRomSerial.ary(1, 104) = "9500"
        
        aryRomSerial.ary(0, 105) = "6be353746310076d1" '9500L0P-412C30E (43 key) s/n 0306563"
        aryRomSerial.ary(0, 106) = "362266e46110076d1" '9500l0p-432c30e  s/n 0317548
        aryRomSerial.ary(0, 107) = "36a266d46010076d1" '9500l0p-432c30e  s/n 0317546 - Found at Baillie Leitchfield KY, currently
            'on loan to
            '
        aryRomSerial.ary(0, 108) = "369266d46f10076d1" '9500l0p-432c30e  s/n 0317547
        aryRomSerial.ary(1, 108) = "9900"
        
        aryRomSerial.ary(0, 109) = "361266d46310076d1" '9500l0p-432c30e  s/n 0317549
        aryRomSerial.ary(0, 110) = "950065049C10076D1" '9500l0p-432c30e  s/n 0317552 Rom serail changed after repair 11/1/2010
        aryRomSerial.ary(1, 110) = "9500" '9500l0p-432c30e  s/n 0317552 Rom serail changed after repair 11/1/2010
        
        aryRomSerial.ary(0, 111) = "36f266d46910076d1" '9500l0p-432c30e  s/n 0317550
        aryRomSerial.ary(0, 112) = "36e266c46e10076d1" '9500l0p-432c30e  s/n 0317513
        aryRomSerial.ary(0, 113) = "36b266c46e10076d1" '9500l0p-432c30e  s/n 0317665
        aryRomSerial.ary(0, 114) = "362266d46110076d1" '9500l0p-432c30e  s/n 0317664
        aryRomSerial.ary(0, 115) = "36f266e46c10076d1" '9500l0p-432c30e  s/n 0317517
        aryRomSerial.ary(0, 116) = "362266e46710076d1" '9500l0p-432c30e  s/n 0317545
        
        aryRomSerial.ary(0, 117) = "0521B5019E4789F11800-0050BF7A60E2" 'HH1307 'Atlanta Hardwoods New Symbol Unit Purchased by Bruce Manalan 08/01/2008
        aryRomSerial.ary(1, 117) = "SYMBOL"
        
        aryRomSerial.ary(0, 118) = "006E604F30137" '9900
        aryRomSerial.ary(1, 118) = "9900"
        
        aryRomSerial.ary(0, 119) = "343326F46410076D1" 'Coastal Spartansburg Chain Tally.  Model#: 9500L0P-422C30E  Serial#: 317566  Manu: 4/9/2008  Install 9/23/2008   HHXXXX
        aryRomSerial.ary(0, 120) = "34F326F46010076D1" 'Coastal Spartansburg Chain Tally.  Model#: 9500L0P-422C30E  Serial#: 317671  Manu: 4/10/2008  Install 9/23/2008  HHXXXX
        aryRomSerial.ary(0, 121) = "91B151646610076D1" 'Walnut Creek Lumber End Tally.  Model#: 9500L0P-121C030E  Serial#: 0275106  Manu: 9/12/2007
        aryRomSerial.ary(0, 122) = "C59736748810076D1" 'Long Range Scanner 9500 Model   Model#: 9501L0P-132C30E  Serial#: 0343049  Manu: 10/13/2008
        aryRomSerial.ary(0, 123) = "347326F46A10076D1" ' 9500 Model   Model#: 9501L0P-422C30E  Serial#: 0317667  Manu: 4/10/2008
        
        aryRomSerial.ary(0, 124) = "36F266C46B10076D1" ' 9500 56 Key Serial 0317518 model 9500L0P-432C30E Purchased 11/5/08 Manu 4/9/08 Probably will be at Panpac
        aryRomSerial.ary(1, 124) = "9900"
        
        aryRomSerial.ary(0, 125) = "366266C46910076D1" ' 9500 56 Key Serial 0317521 model 9500L0P-432C30E Purchased 11/5/08 Manu 4/9/08 Probably will be at Panpac
        aryRomSerial.ary(1, 125) = "9900"
        
        aryRomSerial.ary(0, 126) = "850060A0C031977E0" ' Charles Foreman's Ipaq
        aryRomSerial.ary(1, 126) = "9900"
        
        aryRomSerial.ary(0, 127) = "883651A46810076D1" ' 9500 Model  Serial: 0286095
        aryRomSerial.ary(0, 128) = "34C326F46B10076D1" ' Serial: 0317579
        aryRomSerial.ary(0, 129) = "343326F46A10076D1" ' Serial: 0317580
        aryRomSerial.ary(0, 130) = "368266D46410076D1" ' Serial: 0317598
        aryRomSerial.ary(0, 131) = "34F326F46A10076D1" ' Serial: 0317666
        
        aryRomSerial.ary(0, 132) = "006E604F30144"    ' 9900 43 Key - 'Duplicated on a 35 Key now Unit HH1256
        aryRomSerial.ary(1, 132) = "9900"
        
        aryRomSerial.ary(0, 133) = "1F869901BD48FD21A800-0050BF7A60E2"  'HH1308 Atlanta Hardwoods Symbol
        aryRomSerial.ary(1, 133) = "SYMBOL"
        
        aryRomSerial.ary(0, 134) = "11869901BD4826917800-0050BF7A60E2"  'HH1309 Atlanta Hardwoods Symbol
        aryRomSerial.ary(1, 134) = "SYMBOL"
        
        aryRomSerial.ary(0, 135) = "4B00F4010000-4400300038004700300030003800390039"
        aryRomSerial.ary(1, 135) = "KYMAN" 'DataLogic Kyman
        
        aryRomSerial.ary(0, 136) = "006E604F30138" 'Pan Pac 9900 56 key Serial #: 08020D00
        aryRomSerial.ary(1, 136) = "9900"
        '
        aryRomSerial.ary(0, 137) = "006E604F30136" 'Pan Pac 9900 Serial #: 080331D0056
        aryRomSerial.ary(1, 137) = "9900"
        
        aryRomSerial.ary(0, 138) = "5300F4010000-4400300037004800300033003900370039"
        aryRomSerial.ary(1, 138) = "OTHER" 'Skorpio Datalogic
        
        aryRomSerial.ary(0, 139) = "4D00F4010000-5000300037004600300030003700300032"
        aryRomSerial.ary(1, 139) = "OTHER" 'MEMOR DATALOGIC
        
        
        aryRomSerial.ary(0, 140) = "41A313544910076D1" ' 9500 56 Key Serial: 203249 model 9500L00-131C30E NOT OURS-Purchased by Highway 47 Lumber
        aryRomSerial.ary(1, 140) = "9500"
        
        aryRomSerial.ary(0, 141) = "416313544A10076D1" ' 9500 56 Key Serial 203255 model 9500L00-131C30E NOT OURS-Purchased by Highway 47 Lumber
        aryRomSerial.ary(1, 141) = "9500"
        
        aryRomSerial.ary(0, 142) = "00D2EB5140120" '
        aryRomSerial.ary(1, 142) = "9500"
        
        aryRomSerial.ary(0, 143) = "5B9A78F81C34361DE" ' PocketPC 2003 PDA HP Journada
        aryRomSerial.ary(1, 143) = "OTHER" 'HP Journada


        aryRomSerial.ary(0, 144) = "006E604F30139" 'Serial# 08206D0009  9900L0P-311200
        aryRomSerial.ary(1, 144) = "9900"
        
        aryRomSerial.ary(0, 145) = "006E604F30138" 'Serial# 09115D0008  9900L0P-331200 35 Key
        aryRomSerial.ary(1, 145) = "9900" 'Shit this is another duplicate serial #, need to find another way to license.
        'Same as 136 panpac hh....
        
        aryRomSerial.ary(0, 146) = "006E604F30134" 'Serial# 08331D0044  9900L0P-321200  56 Key
        aryRomSerial.ary(1, 146) = "9900"
        
        aryRomSerial.ary(0, 147) = "8EC46FF44710076D1" 'Serial# 08331D0044  9500B00-121C30
        aryRomSerial.ary(1, 147) = "9500" '- Baillie World Wood Serial # 232147 35 Key 9500
        'Also there 0192408 56 Key 9500 Series handheld, this one appears already licensed for mlIMBS.
        
        aryRomSerial.ary(0, 148) = "91B151646210076D1" 'Serial# 275114 9500 35 Key from Leitchfield Baillie Location
        aryRomSerial.ary(1, 148) = "9500" '- Baillie World Wood Serial # 232147 35 Key 9500
        
        aryRomSerial.ary(0, 149) = "88D651A46D10076D1"
        aryRomSerial.ary(1, 149) = "9500" '??? not sure this is accurate ... Armstrong 43 Key Serial#0286146   HH#1063
        
        
        aryRomSerial.ary(0, 150) = "XXXX"
        aryRomSerial.ary(1, 150) = "HX2" '??? not sure this is accurate ... Armstrong 43 Key Serial#0286146   HH#1063
        
        aryRomSerial.ary(0, 151) = "C57289A48410076D1"
        aryRomSerial.ary(1, 151) = "9500"  'Serial 223857 9500 43 Key logged as HH1239 - It's haessly's loggin handheld held/third/extra one
        
        aryRomSerial.ary(0, 152) = "1341A3010D491361D800-0050BF7A60E2" 'HH1310 Replacement one for originaly one ran over by a forklift :)
        aryRomSerial.ary(1, 152) = "SYMBOL"
                               
        aryRomSerial.ary(0, 153) = "03401FA482A7A60E2" 'HH1311 New Non Triggered Unit added 8/22/09
        aryRomSerial.ary(1, 153) = "SYMBOL2"
        
        aryRomSerial.ary(0, 154) = "22401FA48497A60E2" 'HH1312 New Non Triggered Unit added 8/22/09
        aryRomSerial.ary(1, 154) = "SYMBOL2"
        
        aryRomSerial.ary(0, 155) = "006E604F30133"
        aryRomSerial.ary(1, 155) = "9900" 'New Unit 9/1/2009 - Going to Collins Companies, Boardman Oregon Location.
        
        aryRomSerial.ary(0, 156) = "006E604F30132"
        aryRomSerial.ary(1, 156) = "9900" '56 Key 9900 Serial: 09168D0092  HH1244 Going to Spigelmyer on demo with lumber and logs
        
        
        aryRomSerial.ary(0, 157) = "006E604F30132"
        aryRomSerial.ary(1, 157) = "9900" '56 Key 9900 Serial: 09168D0092  HH1244 Going to Spigelmyer on demo with lumber and logs
        
        aryRomSerial.ary(0, 158) = "006E604F30134" '56 Key 9900 Serial: 08331D0044  HH1244 Going to Spigelmyer on demo with lumber and logs
        aryRomSerial.ary(1, 158) = "9900"
    
        aryRomSerial.ary(0, 159) = "006E604F30144"    ' 9900 35 Key - Duplicated with 43 Key Above
        aryRomSerial.ary(1, 159) = "9900"
    
        aryRomSerial.ary(0, 160) = "693363A45510076D1"    'SN 0269350 ' 9500 From Titusville, setup for Bryan Swift at World Wood.
        aryRomSerial.ary(1, 160) = "9500"
        
        
        aryRomSerial.ary(0, 162) = "19A018249297A60E2" 'New Non Triggered Unit added 1/28/2010
        aryRomSerial.ary(1, 162) = "SYMBOL2"
        
        aryRomSerial.ary(0, 163) = "13A0182496C7A60E2" 'New Non Triggered Unit added 1/28/2010
        aryRomSerial.ary(1, 163) = "SYMBOL2"
        
        
        aryRomSerial.ary(0, 164) = "006E604F30132"    ' 9900 35 Key - Loaner Unit from HHP for tradeshows 1/29/2010
        aryRomSerial.ary(1, 164) = "9900"
        
        aryRomSerial.ary(0, 165) = "006E604F30130"    ' 9900 56 Key - Adam Conway - Superior Hardwoods, Barlow, OH Serial: 9235D0090, 9900L0P-321200
        aryRomSerial.ary(1, 165) = "9900"
        
        aryRomSerial.ary(0, 166) = "006E604F30138" '9900 56 Key HH1319 - Duplicated Serial # for the 3rd Time must find a new way to license devices
        aryRomSerial.ary(1, 166) = "9900"
        
        aryRomSerial.ary(0, 167) = "0056A0E9D0100" 'Intermec Full AlphaNumeric CK3 Serial #302109558085
        aryRomSerial.ary(1, 167) = "CK3ALPHANUMERIC"
        
        aryRomSerial.ary(0, 168) = "8E246ff44710076D1" '35key Baillie
        aryRomSerial.ary(1, 168) = "9500"
        
        
        aryRomSerial.ary(0, 169) = "1F3637302B46D4B14800-0050BF7A60E2" 'SYMBOL AHI - Flooring Waynesboro - Monochrome MC9090 Symbol HH's
        aryRomSerial.ary(1, 169) = "SYMBOL"  'Serial# 7331000502554
        
        aryRomSerial.ary(0, 170) = "162533332846F7911800-0050BF7A60E2" 'SYMBOL AHI - Flooring Waynesboro - Monochrome MC9090 Symbol HH's
        aryRomSerial.ary(1, 170) = "SYMBOL"
        
        aryRomSerial.ary(0, 171) = "0B168400A5466C919800-0050BF7A60E2" 'SYMBOL AHI - Flooring Waynesboro - Monochrome MC9090 Symbol HH's
        aryRomSerial.ary(1, 171) = "SYMBOL"
        
        aryRomSerial.ary(0, 172) = "0F168400A74607D19800-0050BF7A60E2" 'SYMBOL AHI - Flooring Waynesboro - Monochrome MC9090 Symbol HH's
        aryRomSerial.ary(1, 172) = "SYMBOL"
        
                
        aryRomSerial.ary(0, 173) = "006E604F30132" 'Honeywell AH1- Waynesboro - 9900 , Duplicate romserial(006E604F30132)  with (0 , 156)
        aryRomSerial.ary(1, 173) = "9900"
        
        aryRomSerial.ary(0, 174) = "006E604F30131" 'S&J lumber 43-key 9900
        aryRomSerial.ary(1, 174) = "9900"
        
        aryRomSerial.ary(0, 175) = "006E604F30142" 'Augusta 9900 56-key, SN: 10065D00CB
        aryRomSerial.ary(1, 175) = "9900"
        
        aryRomSerial.ary(0, 176) = "006E604F30134" 'Augusta 9900 35-key, SN: 10027D0034
        aryRomSerial.ary(1, 176) = "9900"
        
        
        aryRomSerial.ary(0, 177) = "006E604F30137" 'Augusta 9900 35-key, SN: 10027D0037
        aryRomSerial.ary(1, 177) = "9900"
        
        aryRomSerial.ary(0, 178) = "006E604F30143" 'Augusta - West Point 9900 35-key, SN: 10027D003C
        aryRomSerial.ary(1, 178) = "9900"
        
        aryRomSerial.ary(0, 179) = "006E604F30139" 'Serial# 10065D00C9  9900 56key, duplicate with aryRomSerial.ary(0, 144)
        aryRomSerial.ary(1, 179) = "9900"
        
        aryRomSerial.ary(0, 180) = "006E604F30143" 'Serial# 10176D020C  9900 56key Going to Deveraux as Demo Unit, Returned Most Like Going to emporium now for log system.
        aryRomSerial.ary(1, 180) = "9900"
        
        aryRomSerial.ary(0, 181) = "006E604F30143" 'Serial# 10032D002c 9900 35key Going to Krueger as Demo Unit
        aryRomSerial.ary(1, 181) = "9900"
        
        aryRomSerial.ary(0, 182) = "94406E048C10076D1" ' type: 9500L0P 35key  Desc: Ballie Smyrna ET  S/N: 0270837
        aryRomSerial.ary(1, 182) = "9500"
        
        aryRomSerial.ary(0, 183) = "297524C45E10076D1" 'type: 9500L00 56key desc: Ballie Smyrna S/N: 0252674
        aryRomSerial.ary(1, 183) = "9500"
        
        aryRomSerial.ary(0, 184) = "8E846FF44110076D1" ' type: 9500L00 35key  Desc: Ballie Smyrna ET  S/N: 0244334
        aryRomSerial.ary(1, 184) = "9500"
        
        aryRomSerial.ary(0, 185) = "3A688DF44610076D1" 'type: 9500L00 56key desc: Ballie Smyrna S/N: 0253094
        aryRomSerial.ary(1, 185) = "9500"
        
        aryRomSerial.ary(0, 186) = "3A988DE44710076D1 " 'type: 9500L00 56key desc: Ballie Smyrna S/N: 0253102
        aryRomSerial.ary(1, 186) = "9500"
         
        aryRomSerial.ary(0, 187) = "8E146FF44B10076D1 " 'type: 9500B00 35key desc: Ballie Titus SN 0232319
        aryRomSerial.ary(1, 187) = "9500"
         
        aryRomSerial.ary(0, 188) = "19386EE45010076D1 " 'type: 9500L0P 56key desc: Ballie Titus S/N: 0273385
        aryRomSerial.ary(1, 188) = "9500"
         
        aryRomSerial.ary(0, 189) = "013637302B46B6913800-0050BF7A60E2" 'SYMBOL AHI - Flooring Waynesboro - Monochrome MC9090 Symbol HH's
        aryRomSerial.ary(1, 189) = "SYMBOL"  'Serial# 7331000502554

                                
        aryRomSerial.ary(0, 190) = "91D151646210076D1" 'type: 9500L0P 35key desc: Ballie Leitchfield SN 0275104
        aryRomSerial.ary(1, 190) = "9500"
         
        aryRomSerial.ary(0, 191) = "006E604F30141" 'type: 9500L0P 35key App Hardwoods replacement Serial 10027D003A
        aryRomSerial.ary(1, 191) = "9900"
        
        aryRomSerial.ary(0, 192) = "006E604F30135"  '35 Key 9900 for HH2061 Serial: 10027D0045  9900l0p-331200 DEMO
        aryRomSerial.ary(1, 192) = "9900"
        
        aryRomSerial.ary(0, 193) = "006E604F30145"  '35 Key 9900 for HH2064 12/5/2010
        aryRomSerial.ary(1, 193) = "9900"
        
        aryRomSerial.ary(0, 194) = "DUPLICATED AGAIN"  '35 Key 9900 for HH2066 12/5/2010
        aryRomSerial.ary(1, 194) = "9900"
        
        aryRomSerial.ary(0, 195) = "355286144B10076D1" 'type: 9500 43key desc: eLIMBS
        aryRomSerial.ary(1, 195) = "9500"

        aryRomSerial.ary(0, 196) = "A1D962E4A510076D1" 'type: 9500 43Key desc: eLIMBS
        aryRomSerial.ary(1, 196) = "9500"
        
        aryRomSerial.ary(0, 197) = "006E604F30145" 'type: 9900 56Key desc: Blue Triangle Serial: 10036D00EE
        aryRomSerial.ary(1, 197) = "9900"
        
        aryRomSerial.ary(0, 198) = "006E604F3013810027D003B" 'type: 9900 35Key desc: Blue Triangle Serial: 10027D0038
        aryRomSerial.ary(1, 198) = "9900"

        aryRomSerial.ary(0, 199) = "006E604F30139" 'Duplicate with (0,144), type: 9900 35key Desc: Blue Triangle Serial: 10027D0029
        aryRomSerial.ary(1, 199) = "9900"

        aryRomSerial.ary(0, 200) = "006E604F30138" 'Ddupliate with (0,136), type: 9900 35key Desc: Blue Triangle Serial: 10027D0028
        aryRomSerial.ary(1, 200) = "9900"
        
        aryRomSerial.ary(0, 201) = "006E604F30142HH2100" 'type: 9900 56key Desc: Blue Triangle Serial: 10177D01FB
        aryRomSerial.ary(1, 201) = "9900"          'HH2100
        
        aryRomSerial.ary(0, 202) = "006E604F30134" 'type: 9900 35key Desc: Serial: 10027D0034
        aryRomSerial.ary(1, 202) = "9900"
        
        aryRomSerial.ary(0, 203) = "006E604F30139" 'type: 9900 35key desc: Serial: 10269D0269
        aryRomSerial.ary(1, 203) = "9900"

        aryRomSerial.ary(0, 203) = "006E604F30134HH2081" 'type: 9900 56key desc: Serial: 10177D0284
        aryRomSerial.ary(1, 203) = "9900"
        
        aryRomSerial.ary(0, 203) = "006E604F30134HH2081" 'type: 9900 56key desc: Serial: 10177D0284
        aryRomSerial.ary(1, 203) = "9900"
        
        aryRomSerial.ary(0, 204) = "0A1001D43410076D1" 'type: 9500 Bruce Manalan acquired in purchase of craig lumber
        aryRomSerial.ary(1, 204) = "9500"
        
        aryRomSerial.ary(0, 205) = "05400D442F10076D1" 'type: 9500 Bruce Manalan acquired in purchase of craig lumber
        aryRomSerial.ary(1, 205) = "9500"
        

        aryRomSerial.ary(0, 206) = "006E604F3013408206D006D" 'Serial# 08206D006D  9900L0P-321200  56 Key Corsica Matson Receiving
        aryRomSerial.ary(1, 206) = "9900"
        
        
        '*************TAYLOR LUMBER - LUMBER HANDHELDS
        aryRomSerial.ary(0, 207) = "006E604F30138HH2066" 'type: 9900 35 KEY TAYLOR LUMBER HH2066
        aryRomSerial.ary(1, 207) = "9900"
        aryRomSerial.ary(2, 207) = "IPSM" 'New Featuer to Determine the Install Path
        
        aryRomSerial.ary(0, 208) = "006E604F30143HH2102" 'type: 9900 35 KEY TAYLOR LUMBER HH2102
        aryRomSerial.ary(1, 208) = "9900"
        aryRomSerial.ary(2, 208) = "IPSM" 'New Featuer to Determine the Install Path
        
        aryRomSerial.ary(0, 209) = "006E604F30142HH2098" 'type: 9900 35 KEY TAYLOR LUMBER HH2098
        aryRomSerial.ary(1, 209) = "9900"
        aryRomSerial.ary(2, 209) = "IPSM" 'New Featuer to Determine the Install Path
        
        aryRomSerial.ary(0, 210) = "006E604F30131HH2099" 'type: 9900 35 KEY TAYLOR LUMBER HH2099
        aryRomSerial.ary(1, 210) = "9900"
        aryRomSerial.ary(2, 210) = "IPSM" 'New Featuer to Determine the Install Path
        
        
        '*************TAYLOR LUMBER - LUMBER HANDHELDS
        
        aryRomSerial.ary(0, 211) = "19586EE45C10076D1" '35 Key 9500 bAILLIE SMYRNA hh1039 serial 0270838
        aryRomSerial.ary(1, 211) = "9500"
        
        '*************GUTCHESS - LUMBER HANDHELDS
        'All 99EX have the same ROM serial, GUTCHESS has 8 handhelds.
        aryRomSerial.ary(0, 212) = "50F0063006B000000HH2131" '34 key 99EX
        aryRomSerial.ary(1, 212) = "7600" 'Actually a 99EX
                
        
        aryRomSerial.ary(0, 212) = "50F0063006B000000HH2131" '54 key 99EX
        aryRomSerial.ary(1, 212) = "7600" 'Actually a 99EX
        
        aryRomSerial.ary(0, 213) = "006E604F30146" '35 Key 9900 HH2154 Serial# 11094D00AF
        aryRomSerial.ary(1, 213) = "9900"
        
        


        aryRomSerial.ary(0, 214) = "006E604F30131HHXXX" 'type: 9900 35 KEY  Serial 09237D00F1  Part#  9900-LOP-331200  9900 35 Key
        aryRomSerial.ary(1, 214) = "9900"

        aryRomSerial.ary(0, 215) = "006E604F30131HHXXX" 'type: 9900 35 KEY Serial Serial 09228D00BE Part#  9900-LOP-331200  9900 35 Key
        aryRomSerial.ary(1, 215) = "9900"

        aryRomSerial.ary(0, 216) = "006E604F30131HH2155" 'type: 9900 35 KEY TAYLOR Serial 09235D005A Part#  9900-LOP-331200  9900 35 Key
        aryRomSerial.ary(1, 216) = "9900"

        aryRomSerial.ary(0, 217) = "403663344B10076D1" 'type: 9500 56 KEY TAYLOR Serial 0213900 New River 56 Key trade license
        aryRomSerial.ary(1, 217) = "9500"

        aryRomSerial.ary(0, 218) = "7D0523244B10076D1" 'type: 9500 56 KEY TAYLOR Serial 0186883 New River 56 Key trade license
        aryRomSerial.ary(1, 218) = "9500"


        'All 99EX have the same ROM serial, GUTCHESS has 8 handhelds.
        aryRomSerial.ary(0, 220) = "50F0063006B00000011118D0079" '2137 34 Key Gutchess 99EX
        aryRomSerial.ary(1, 220) = "7600" 'Actually a 99EX


        'All 99EX have the same ROM serial,
        aryRomSerial.ary(0, 221) = "50F0063006B000000HH2165" '2165 34 Key Gutchess 99EX Serial# 11153DO26F
        aryRomSerial.ary(1, 221) = "7600" 'Actually a 99EX - Going to Turman in Virginia


        'Alot of 9900's have the same ROM serial,
        aryRomSerial.ary(0, 222) = "006E604F3014310032D002C" ' 34 Key Krueger - Serial # 10032D002C
        aryRomSerial.ary(1, 222) = "9900" 'Sent as a loaner to AHI on 11/16/11
        

        'Small 6100 Series Honeywell, probably going to NHLA for now
        '****took this out because somebody duplicated it below and screwed up model # so doesn't work now
'''        aryRomSerial.ary(0, 223) = "570049004E00430045003500300030000000-3130303232333030313900" ' 6100 Series demo unit
'''        aryRomSerial.ary(1, 223) = "7600"


        aryRomSerial.ary(0, 224) = "006E604F30143HH2035" 'Augusta - North Garden
        aryRomSerial.ary(1, 224) = "9900"

        
        'All 99EX have the same ROM serial, GUTCHESS has 8 handhelds.
        aryRomSerial.ary(0, 225) = "50F0063006B00000011119D0083" '2147 34 Key Gutchess 99EX
        aryRomSerial.ary(1, 225) = "7600" 'Actually a 99EX
        
        aryRomSerial.ary(0, 226) = "50F0063006B00000011112D0111" '2181 55 Key Gutchess 99EX
        aryRomSerial.ary(1, 226) = "7600" 'Actually a 99EX
        
        aryRomSerial.ary(0, 227) = "570049004E00430045003500300030000000-3039333630333030313700" '2181 55 Key Gutchess 99EX
        aryRomSerial.ary(1, 227) = "7600" 'Actually a 6000 OR SOMETHING DEMO UNIT USES HONEYWELL NOT IPSM PATH
        
        'All 99EX have the same ROM serial, GUTCHESS has 8 handhelds.
        aryRomSerial.ary(0, 228) = "50F0063006B00000011127D0005" 'Key Gutchess 99EX
        aryRomSerial.ary(1, 228) = "7600" 'Actually a 99EX
        
        aryRomSerial.ary(0, 229) = "460061006c0063006f006e00580033000000-44313142303938313200"
        aryRomSerial.ary(1, 229) = "Falcon" 'DataLogic Kyman
        
        aryRomSerial.ary(0, 230) = "006E604F30143HH2048" ' 9900 Westpoint AHI
        aryRomSerial.ary(1, 230) = "9900" 'DataLogic Kyman
        
         
        aryRomSerial.ary(0, 232) = "50F0063006B00000011126D006C" 'HH2138 Gutchess Lumber Cortland
        aryRomSerial.ary(1, 232) = "7600" '99EX
        
        aryRomSerial.ary(0, 233) = "50F0063006B00000011119D00AC" ' HH2139 Gutchess Lumber Cortland"
        aryRomSerial.ary(1, 233) = "7600" '99EX
        
       aryRomSerial.ary(0, 234) = "50F0063006B00000011142D00EE" 'HH2160 Serial: 11142D00EE eLIMBS Demo Handheld 55 Key 99EXLW3-GC211Xe 11/7/11
       aryRomSerial.ary(1, 234) = "7600" '99EX
       
       aryRomSerial.ary(0, 235) = "50F0063006B00000011230D014A" 'HH2170 Serial: 11230D014A eLIMBS Demo Handheld 34 Key 99EXL01-0C212SE eLIMBS Demo End Tally Currently 11/7/11
       aryRomSerial.ary(1, 235) = "7600" '99EX
       
        aryRomSerial.ary(0, 236) = "880651A46010076D1"
        aryRomSerial.ary(1, 236) = "9500" '- Loaner Handheld at Clendenin using for Chain Tally 43 Key...
        
        aryRomSerial.ary(0, 237) = "6B3353746110076D1"
        aryRomSerial.ary(1, 237) = "9500" '- Loaner Handheld going to College Hill Lumber 12/23/2011 - 43 Key
        
        aryRomSerial.ary(0, 238) = "50F0063006B00000011236D038F"
        aryRomSerial.ary(1, 238) = "7600" '- 34 Key 99EX Anderson Tully HH2174
        
        aryRomSerial.ary(0, 239) = "50F0063006B00000011237D011"
        aryRomSerial.ary(1, 239) = "7600" '- 34 Key 99EX Anderson Tully HH2173
        
        aryRomSerial.ary(0, 240) = "50F0063006B00000011230D00F8"
        aryRomSerial.ary(1, 240) = "7600" '- 55 Key 99EX Anderson Tully HH2175
        
        aryRomSerial.ary(0, 241) = "50F0063006B00000011179D03FF"
        aryRomSerial.ary(1, 241) = "7600" '- 55 Key 99EX Anderson Tully HH2176
        
        aryRomSerial.ary(0, 242) = "50F0063006B00000011179D0129"
        aryRomSerial.ary(1, 242) = "7600" '- 55 Key 99EX Anderson Tully HH2177
        
        aryRomSerial.ary(0, 243) = "50F0063006B00000011230D0108"
        aryRomSerial.ary(1, 243) = "7600" '- 34 Key 99EX Anderson Tully HH2178
        
        aryRomSerial.ary(0, 244) = "50F0063006B00000011112D010B"
        aryRomSerial.ary(1, 244) = "7600" '- 55 Key 99EX Gutchess HH2121
        
        aryRomSerial.ary(0, 245) = "50F0063006B00000011112D0119"
        aryRomSerial.ary(1, 245) = "7600" '- 55 Key 99EX Gutchess HH2122
        
        aryRomSerial.ary(0, 246) = "50F0063006B00000011118D00AB"
        aryRomSerial.ary(1, 246) = "7600" '- 55 Key 99EX Gutchess HH2136
        
        aryRomSerial.ary(0, 247) = "50F0063006B00000011113D0031"
        aryRomSerial.ary(1, 247) = "7600" '- 55 Key 99EX Gutchess HH2119
        
        aryRomSerial.ary(0, 248) = "50F0063006B00000011112D0108"
        aryRomSerial.ary(1, 248) = "7600" '- GC211XE Key 99EX Gutchess HH2120

        aryRomSerial.ary(0, 249) = "50F0063006B000000"
        aryRomSerial.ary(1, 249) = "7600" '- 55 Key 99EX Gutchess HH2122
        
        
        aryRomSerial.ary(0, 250) = "50F0063006B00000011194D01C3"  'HH2193
        aryRomSerial.ary(1, 250) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 251) = "50F0063006B00000011194D02A8"  'HH2197
        aryRomSerial.ary(1, 251) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 252) = "50F0063006B00000011194D0199"  'HH2191
        aryRomSerial.ary(1, 252) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 253) = "50F0063006B00000011200D00B1"  'HH2189
        aryRomSerial.ary(1, 253) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 254) = "50F0063006B00000011201D017B"  'HH2200
        aryRomSerial.ary(1, 254) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 255) = "50F0063006B00000011195D0035"  'HH2199
        aryRomSerial.ary(1, 255) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 256) = "50F0063006B00000011194D0293"  'HH2196
        aryRomSerial.ary(1, 256) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 257) = "50F0063006B00000011200D00EE"  'HH2190
        aryRomSerial.ary(1, 257) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 258) = "50F0063006B00000011194D02B3"  'HH2198
        aryRomSerial.ary(1, 258) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 259) = "50F0063006B00000011194D01C2"  'HH2192
        aryRomSerial.ary(1, 259) = "7600" '- 99EX Gutchess
        
        aryRomSerial.ary(0, 260) = "50F0063006B00000012189D011C"  'HH2287 Anderson Tully Vix
        aryRomSerial.ary(1, 260) = "7600" '-
        
        aryRomSerial.ary(0, 261) = "50F0063006B00000012113D0177"  'HH2283 Anderson Tully Vix - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 261) = "7600" '-
        
        aryRomSerial.ary(0, 262) = "50F0063006B000000HH3002"  'HH3002 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 262) = "7600" '-
        
        aryRomSerial.ary(0, 263) = "50F0063006B000000HH3003"  'HH3003 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 263) = "7600" '-
        
        aryRomSerial.ary(0, 264) = "50F0063006B000000HH3004"  'HH3004 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 264) = "7600" '-
        
        aryRomSerial.ary(0, 265) = "50F0063006B000000HH3005"  'HH3005 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 265) = "7600" '-
        
        aryRomSerial.ary(0, 266) = "50F0063006B000000HH3006"  'HH3006 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 266) = "7600" '-
        
        aryRomSerial.ary(0, 267) = "50F0063006B000000HH3007"  'HH3007 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 267) = "7600" '-
        
        aryRomSerial.ary(0, 268) = "50F0063006B000000HH3008"  'HH3008 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 268) = "7600" '-
        
        aryRomSerial.ary(0, 269) = "50F0063006B000000HH3009"  'HH3009 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 269) = "7600" '-
        
        aryRomSerial.ary(0, 270) = "50F0063006B000000HH3010"  'HH3010 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 270) = "7600" '-
        
        aryRomSerial.ary(0, 271) = "50F0063006B000000HH3011"  'HH3011 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 271) = "7600" '-
        
        aryRomSerial.ary(0, 272) = "50F0063006B000000HH3012"  'HH3012 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 272) = "7600" '-
        
        aryRomSerial.ary(0, 273) = "50F0063006B000000HH3013"  'HH3013 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 273) = "7600" '-
        
        aryRomSerial.ary(0, 274) = "50F0063006B000000HH3014"  'HH3014 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 274) = "7600" '-
        
        aryRomSerial.ary(0, 275) = "50F0063006B000000HH3015"  'HH3015 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 275) = "7600" '-
        
        aryRomSerial.ary(0, 276) = "50F0063006B000000HH3016"  'HH3016 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 276) = "7600" '-
        
        aryRomSerial.ary(0, 277) = "50F0063006B000000HH3017"  'HH3017 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 277) = "7600" '-
        
        aryRomSerial.ary(0, 278) = "50F0063006B000000HH3018"  'HH3018 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 278) = "7600" '-
        
        aryRomSerial.ary(0, 279) = "50F0063006B000000HH3019"  'HH3019 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 279) = "7600" '-
        
        aryRomSerial.ary(0, 280) = "50F0063006B000000HH3020"  'HH3020 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 280) = "7600" '-
        
        aryRomSerial.ary(0, 281) = "50F0063006B000000HH3021"  'HH3021 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 281) = "7600" '-
        
        aryRomSerial.ary(0, 282) = "50F0063006B000000HH3022"  'HH3022 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 282) = "7600" '-
        
        aryRomSerial.ary(0, 283) = "50F0063006B000000HH3023"  'HH3023 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 283) = "7600" '-
        
        aryRomSerial.ary(0, 284) = "50F0063006B000000HH3024"  'HH3024 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 284) = "7600" '-
        
        aryRomSerial.ary(0, 285) = "50F0063006B000000HH3025"  'HH3025 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 285) = "7600" '-
        
        aryRomSerial.ary(0, 286) = "50F0063006B000000HH3026"  'HH3026 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 286) = "7600" '-
        
        aryRomSerial.ary(0, 287) = "50F0063006B000000HH3027"  'HH3027 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 287) = "7600" '-
        
        aryRomSerial.ary(0, 288) = "50F0063006B000000HH3028"  'HH3028 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 288) = "7600" '-
        
        aryRomSerial.ary(0, 289) = "50F0063006B000000HH3029"  'HH3029 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 289) = "7600" '-
        
        aryRomSerial.ary(0, 290) = "50F0063006B000000HH3030"  'HH3030 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 290) = "7600" '-
        
        aryRomSerial.ary(0, 291) = "50F0063006B000000HH3031"  'HH3031 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 291) = "7600" '-
        
        aryRomSerial.ary(0, 292) = "50F0063006B000000HH3032"  'HH3032 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 292) = "7600" '-
        
        aryRomSerial.ary(0, 293) = "50F0063006B000000HH3033"  'HH3033 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 293) = "7600" '-
        
        aryRomSerial.ary(0, 294) = "50F0063006B000000HH3034"  'HH3034 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 294) = "7600" '-
        
        aryRomSerial.ary(0, 295) = "50F0063006B000000HH3035"  'HH3035 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 295) = "7600" '-
        
        aryRomSerial.ary(0, 296) = "50F0063006B000000HH3036"  'HH3036 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 296) = "7600" '-
        
        aryRomSerial.ary(0, 297) = "50F0063006B000000HH3037"  'HH3037 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 297) = "7600" '-
        
        aryRomSerial.ary(0, 298) = "50F0063006B000000HH3038"  'HH3038 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 298) = "7600" '-
        
        aryRomSerial.ary(0, 299) = "50F0063006B000000HH3039"  'HH3039 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 299) = "7600" '-
        
        aryRomSerial.ary(0, 300) = "0521B5019E4789F118000050BF7A60E2" 'HH2191 'Atlanta Hardwoods New Symbol Unit Purchased by Bruce Manalan 08/01/2008
        aryRomSerial.ary(1, 300) = "SYMBOL"
        
        aryRomSerial.ary(0, 301) = "50F0063006B00000012113D033DF" 'HH2298 'Anderson Tully - Inspection of Graders and some Purch Loads - Added 1/31/13
        aryRomSerial.ary(1, 301) = "7600"
        
        aryRomSerial.ary(0, 302) = "796925344A10076D1" 'HIGHWAY 47
        aryRomSerial.ary(1, 302) = "9500"
        
        aryRomSerial.ary(0, 303) = "50F0063006B000000HH6000"  'HH3039 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 303) = "7600" '-
        
        aryRomSerial.ary(0, 304) = "01B019547117A60E2"  'HH3039 - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 304) = "SYMBOL2" '-
        
        aryRomSerial.ary(0, 305) = "50F0063006B00000013070D8052"  'HH2320 Serial 13070D8052 Anderson Tully Shipping 6/28/13 for Geoff/Receiving - 2 Handhelds - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 305) = "7600" '-
        
        aryRomSerial.ary(0, 306) = "50F0063006B00000013070D8050"  'HH2319 Serial Num: 13070D8050 Anderson Tully Shipping 6/28/13 for Geoff/Receiving - 2 Handhelds - place holder in case hh dev gets messed up and cant'license new units
        aryRomSerial.ary(1, 306) = "7600" '-
                                
        aryRomSerial.ary(0, 307) = "50F0063006B00000013070D805E" 'HH2322 Serial Num:  13070D805E Taylor Lumber Company 99EX 34 Key 7/12/13 - Chain Tally
        ''*****(Because of the older version, this one was shipped with license (serial# 11119D0083) that met the necessary date range of their mLIMBS.Arm.Cab - This is the corret license Info for the handheld
        aryRomSerial.ary(1, 307) = "7600" '-
        
        aryRomSerial.ary(0, 308) = "50F0063006B000000HH2323"  'HH2323 Serial Num: 13070D806C Taylor Lumber Company 99EX 34 Key 7/12/13 - Lumber We Think...
        ''*****(Because of the older version, this one was shipped with license (serial# 11112D0111) that met the necessary date range of their mLIMBS.Arm.Cab - This is the corret license Info for the handheld
        aryRomSerial.ary(1, 308) = "7600" '-
        
        
        aryRomSerial.ary(0, 309) = "50F0063006B00000012189D011D"  'HH2291 Serial Num: 12189D011D Andersontully RMA Return had wrong ID originally.
        aryRomSerial.ary(1, 309) = "7600" '-
        
        aryRomSerial.ary(0, 310) = "50F0063006B00000013086D826E"  'HH7327 Serial Num: 13086D826E 55 Key St James HH (Baillie 99EX)
        aryRomSerial.ary(1, 310) = "7600" '-
        
        
        aryRomSerial.ary(0, 311) = "50F0063006B00000013070D8073"  'HH7328 Serial Num: 13070D8073 Beard Handheld
        aryRomSerial.ary(1, 311) = "7600" '-
                
        aryRomSerial.ary(0, 312) = "50F0063006B00000013070D806B"  'HH2324 - Baillie Worldwood Air dried / stacker HH paired with BT P4T
        aryRomSerial.ary(1, 312) = "7600" '-
                
        aryRomSerial.ary(0, 313) = "50F0063006B00000011230D00BB"  'HH2233 - T&D Thompson 11230D00BB - Wasn't originally licnesed, completed during an RMA on 8/29/13
        aryRomSerial.ary(1, 313) = "7600" '-
        
        '****FOUND on 9/4/13 - Older handheld, but never licesned before currently at Matson.
        aryRomSerial.ary(0, 314) = "50F0063006B00000012113D0335"   'HH2297 @ Serial 12113D0335 - Matson Lumber Company (For a while, no record of it ever being licensed..or how it got there..but as of 9/4/13 it was definitely there!
        ''and it is still there as of 10/14/13 - used by Brad (Grader...not end tallier)
        aryRomSerial.ary(1, 314) = "7600" '-
                            


        'Alot of 9900's have the same ROM serial,
        aryRomSerial.ary(0, 316) = "006E604F3014110058D008A" ' Brian Wilson Demo Lumber on one of his log handhelds
        aryRomSerial.ary(1, 316) = "9900" 'Loaded by PCC on 9/17/2013
        
        
        aryRomSerial.ary(0, 317) = "50F0063006B00000013150D80C7"  'HH7335 - SErial 13150D80C7 - Matson LUmber Company with 5 year simple service Installed/delieverd by PCC/Brewster on 10/15/13
        aryRomSerial.ary(1, 317) = "7600" '-
        
        aryRomSerial.ary(0, 318) = "50F0063006B00000011230D0163"  'HH2231 Serial 11230D0163    99EXL01-0C212SE - At Matson lumber, been there a while, Brads ET handheld...not sure why not licensed before!!
        aryRomSerial.ary(1, 318) = "7600" '-
        
        aryRomSerial.ary(0, 319) = "50F0063006B00000013113D821D"  'HH2317 Serial 113113D821D    99EXL01-0C212SE - At Fredonia Forest Products, been there a while,not sure why not licensed before!!
        aryRomSerial.ary(1, 319) = "7600" '-
        
        aryRomSerial.ary(0, 320) = "50F0063006B00000012128D0166"  'HH2307 Serial 12128D0166    99EXL01-0C212SE - At BPM lumber to replace the one they "think has issues"
        aryRomSerial.ary(1, 320) = "7600" '-
        
        aryRomSerial.ary(0, 321) = "50F0063006B00000013150D80D2"  'HH7336 - SErial 13150D80DZ - French Fuckers
        aryRomSerial.ary(1, 321) = "7600" '-
        
        aryRomSerial.ary(0, 322) = "50F0063006B00000012189D0524"  'HH2293 - SErial
        aryRomSerial.ary(1, 322) = "7600" '-
        
        aryRomSerial.ary(0, 323) = "19432E345D87A60E2" 'Symbol MC9090 Purchased by Bruce Manalan for Atlanta Hardwoods...I assume it's a non-triggered unit.
        aryRomSerial.ary(1, 323) = "SYMBOL2"
        

        aryRomSerial.ary(0, 324) = "50F0063006B00000013151D8099"  'HH7342 Serial 13151D8099    99EXL01-0C212SE - Aacer Flooring mLIMBS
        aryRomSerial.ary(1, 324) = "7600" '-
        
        aryRomSerial.ary(0, 325) = "50F0063006B00000013154D8105"  'HH7341 Serial 13154D8105    99EXL01-0C212SE - Aacer Flooring mLIMBS
        aryRomSerial.ary(1, 325) = "7600" '-
        
        aryRomSerial.ary(0, 326) = "50F0063006B00000013151D81A4"  'HH7343 Serial 13151D81A4    99EXL01-0C212SE - Aacer Flooring mLIMBS
        aryRomSerial.ary(1, 326) = "7600" '-


        aryRomSerial.ary(0, 327) = "1E60199474C7A60E2" 'Symbol MC9090 Purchased by Bruce Manalan for Atlanta Hardwoods...I assume it's a non-triggered unit.
        aryRomSerial.ary(1, 327) = "SYMBOL2"
        
        aryRomSerial.ary(0, 328) = "50F0063006B00000013151D8072"  'HH7340 Serial 13151D8072    99EXL01-0C212SE - BYLC MLIMBS
        aryRomSerial.ary(1, 328) = "7600" '-
        
        
        aryRomSerial.ary(0, 329) = "50F0063006B00000013151D8001"  'HH7347 Serial 13151D8001    99EXL01-0C212SE - BYLC MLIMBS
        aryRomSerial.ary(1, 329) = "7600" '-
        
        aryRomSerial.ary(0, 330) = "50F0063006B00000012351D83A3"  'HHL2310 Serial 12351D83A3    99EX - loaner MLIMBS
        aryRomSerial.ary(1, 330) = "7600" '-
        
        aryRomSerial.ary(0, 331) = "50F0063006B00000012188D028E"  'HHL2286 Serial 12188D028E    99EX- Loaner MLIMBS
        aryRomSerial.ary(1, 331) = "7600" '-
                
        aryRomSerial.ary(0, 332) = "50F0063006B00000013150D8228"  'HH7353 Serial 13150D8228    99EX- Atlanta Hardwoods - Bruce Manalan New Handheld for Distribution
        aryRomSerial.ary(1, 332) = "7600" '-
                
        aryRomSerial.ary(0, 333) = "50F0063006B00000013150D8163"  'HH7352 Serial 13150D8163    99EX- Atlanta Hardwoods - Bruce Manalan Demo/New Handheld  for Distribution
        aryRomSerial.ary(1, 333) = "7600" '-
        
        aryRomSerial.ary(0, 334) = "50F0063006B00000013086D805F"  'HH7351 Serial 13086D805F    99EX- Atlanta Hardwoods - Bruce Manalan Demo/New Handheld  for Distribution 2/11/2014
        aryRomSerial.ary(1, 334) = "7600" '-
        
        aryRomSerial.ary(0, 335) = "50F0063006B00000013150D80D1"  'HH7350 Serial 13150D80D1    99EX- Cummings - 1 of 4 Sending Early for Chain Tally to Help a Brother Out 2/12/13
        aryRomSerial.ary(1, 335) = "7600" '-
        
        
        aryRomSerial.ary(0, 336) = "50F0063006B00000013362D8074"  'HH Serial  13362D8074  HH7365 55KEY FREDONIA   99EX-
        aryRomSerial.ary(1, 336) = "7600" '-
        
        aryRomSerial.ary(0, 337) = "50F0063006B00000014019D830D"  'HH Serial 14019D830D  HH7364 55 KEY SUGAR GROVE    99EX-
        aryRomSerial.ary(1, 337) = "7600" '-
        
        aryRomSerial.ary(0, 338) = "50F0063006B00000013086D8084"  'HH Serial 13086D8084 HH7362     99EX- 35 KEY - Pre Licensed - Not Assigned Yet
        aryRomSerial.ary(1, 338) = "7600" '-
        
        aryRomSerial.ary(0, 339) = "50F0063006B00000013085D837E"  'HH Serial 13085D837E HH7363   99EX- 35 KEY - Pre Licensed - Not Assigned Yet
        aryRomSerial.ary(1, 339) = "7600" '-
        
        
        aryRomSerial.ary(0, 340) = "50F0063006B00000013086D8095" '13086D8095 - HH7355 Cummings Lumber Handheld"
        aryRomSerial.ary(1, 340) = "7600" '-
        
        aryRomSerial.ary(0, 341) = "50F0063006B00000013086D8059" '13086D8059-  HH7357  Cummings Lumber Handheld"
        aryRomSerial.ary(1, 341) = "7600" '-
        
        aryRomSerial.ary(0, 342) = "50F0063006B00000013086D8093" '13086D8093- HH7356  Cummings Lumber Handheld
        aryRomSerial.ary(1, 342) = "7600" '-

        aryRomSerial.ary(0, 343) = "50F0063006B00000013086D8061" '13086D8061- HH7367  AHI North Garden - Ole Handheld
        aryRomSerial.ary(1, 343) = "7600" '-

        aryRomSerial.ary(0, 344) = "50F0063006B00000013086D800F" '13086D800F- HH7368  Endeavor Hardwoods, LLC Per Ben - Chain Tally
        aryRomSerial.ary(1, 344) = "7600" '-
        
        aryRomSerial.ary(0, 345) = "50F0063006B00000013086D805E" '13086D805E- HH7359 SN – 13086D805E mLIMBS and mLOGS Baxter Lumber
        aryRomSerial.ary(1, 345) = "7600" '-
                

        aryRomSerial.ary(0, 346) = "8E746FF44010076D1" ' type: 9500 Temporary Licenese For Aaron Fouts Handheld While his is repaired
        aryRomSerial.ary(1, 346) = "9500"
        
        
        aryRomSerial.ary(0, 347) = "50F0063006B00000014105D8185" ' 99EX Serial# 14105D8185   HH# 7370     Client# Baillie Smyrna -99EX 35Key
        aryRomSerial.ary(1, 347) = "7600"
        
        aryRomSerial.ary(0, 348) = "50F0063006B00000014106D829D" ' 99EX Serial# 14106D829D   HH# 7371    Client# Baillie Smyrna -99EX 35Key
        aryRomSerial.ary(1, 348) = "7600"
        
        aryRomSerial.ary(0, 349) = "50F0063006B00000014105D8027" ' 99EX Serial# 14105D8027   HH#  7372    Client# Baillie Smyrna -99EX 35Key
        aryRomSerial.ary(1, 349) = "7600"
        
        aryRomSerial.ary(0, 350) = "50F0063006B00000014106D8210" ' 99EX Serial# 14106D8210   HH# 7373    Client# Baillie Smyrna -99EX 35Key
        aryRomSerial.ary(1, 350) = "7600"
        
        aryRomSerial.ary(0, 351) = "50F0063006B00000014105D8053" ' 99EX Serial# 14105D8053    HH# 7374    Client# Lester Yoder - Independent Grader -99EX 35Key
        aryRomSerial.ary(1, 351) = "7600"
        
        aryRomSerial.ary(0, 352) = "50F0063006B00000014105D801F" ' 99EX Serial# 14105D801F   HH# 7382     Client# Northeastern States Lumber  -99EX 35Key
        aryRomSerial.ary(1, 352) = "7600"

        aryRomSerial.ary(0, 353) = "50F0063006B00000014116D801F" ' 99EX Serial# 14116D801F   HH# 7384     Client# Matson Lumber -99EX 35Key
        aryRomSerial.ary(1, 353) = "7600"


        aryRomSerial.ary(0, 354) = "50F0063006B00000014114D80A1" ' 99EX Serial# 14114D80A1   HH# 7385  Client# Turman Lumber -99EX 35Key - Chain Tally is what it is planned to be used for
        aryRomSerial.ary(1, 354) = "7600"


        '*************Dana Spessert NHLA - This handheld was already licensed, and this entry screwed it up (whoever added it) ..6100 is an invalid model
        'within the code
      ''***BAD ENTRY'  aryRomSerial.ary(0, 355) = "570049004e00430045003500300030000000-3130303232333030313900" '6100
      ''***BAD ENTRY'  aryRomSerial.ary(1, 355) = "6100"
        
        'Updated by PCC****2/8/2017 for Baillie/Paul Hare 6100 series HH Use
        aryRomSerial.ary(0, 355) = "570049004e00430045003500300030000000-3130303232333030313900" '6100
        aryRomSerial.ary(1, 355) = "7600"

        aryRomSerial.ary(0, 356) = "50F0063006B00000014105D8197" ' 99EX Serial# 14105D8197   HH# 7388  Horizon 99EX 35 Key
        aryRomSerial.ary(1, 356) = "7600"
        
'        aryRomSerial.ary(0, ) = "50F0063006B000000" ' 99EX Serial#    HH#     Client#
'        aryRomSerial.ary(1, ) = "7600"
        
        aryRomSerial.ary(0, 357) = "50F0063006B00000014186D82EF" ' 99EX Serial# 14186D82EF   HH# 7400  Mullican 99EX 35 Key  99EXL01-0C212SE 9/2/14 End Tally - Next Day - Last Minute Setup and Shipment.
        aryRomSerial.ary(1, 357) = "7600"
        
        aryRomSerial.ary(0, 358) = "50F0063006B00000014113D8035" ' 99EX Serial# 14113D8035   HH# 7410  Stoney Point 99EX 35 Key  99EXL01-0C212SE 9/17/14 Green Chain Tally.
        aryRomSerial.ary(1, 358) = "7600"
        
        
        aryRomSerial.ary(0, 359) = "50F0063006B00000014186D8297" ' 99EX Serial# 14186D8297   HH# 7411  Salem Hardwood mLIMBS 9/23/2014 99EX 35 Key  99EXL01-0C212SE
        aryRomSerial.ary(1, 359) = "7600"
        
        aryRomSerial.ary(0, 360) = "50F0063006B00000011231D0062" ' 99EX Serial# 11231D0062   HH# HH2232  Matson Lumber  - Old 2012 Handheld that wasn't licensed due to version not being current.  Licensing now
        aryRomSerial.ary(1, 360) = "7600"
        
        
        
        aryRomSerial.ary(0, 361) = "50F0063006B00000014186D8286" ' 99EX Serial# 14186D8286   HH# 7413     Client# Gilkey Lumber -99EX 35Key
        aryRomSerial.ary(1, 361) = "7600"
        
        aryRomSerial.ary(0, 362) = "50F0063006B00000014186D82DD" ' 99EX Serial# 14186D82DD   HH# 7401     Client# Cherry Hill Lumber  -99EX 35Key
        aryRomSerial.ary(1, 362) = "7600"
        

        aryRomSerial.ary(0, 363) = "006E604F3013310065D0113"
        aryRomSerial.ary(1, 363) = "9900" 'Sugar Grove Added after the fact, previously wasn't licensed it seems - HH2094 Serial 10065D0113
        
        
        aryRomSerial.ary(0, 364) = "50F0063006B00000014116D8021" ' 99EX Serial# 14116D8021   HH# 7395     Client# Cherry Hill Lumber  -99EX 35Key
        aryRomSerial.ary(1, 364) = "7600"
        
        aryRomSerial.ary(0, 365) = "50F0063006B00000014258D8041"  'HH Serial 14258D8041 HH7428 34 Key Log & Lumber Licensed - Gilkey Lumber 11/17/14
        aryRomSerial.ary(1, 365) = "7600"
        
        aryRomSerial.ary(0, 366) = "50F0063006B00000014258D80EA" ' 99EX Serial# 14258D80EA   HH# 7424     Client# McKay Hardwoods -99EX 34Key
        aryRomSerial.ary(1, 366) = "7600"
        
        aryRomSerial.ary(0, 367) = "50F0063006B00000011230D0089" ' 99EX Serial# 11230D0089   HH# 2230     Client# CC Cook & Son -99EX 34Key
        aryRomSerial.ary(1, 367) = "7600"
        
        aryRomSerial.ary(0, 368) = "50F0063006B00000012128D0166" ' 99EX Serial# 12128D0166   HH# 2307     Client# Thompson Hardwoods - 99EX 55Key
        aryRomSerial.ary(1, 368) = "7600"

        aryRomSerial.ary(0, 369) = "50F0063006B00000014258D8038" ' 99EX Serial# 14258D8038   HH# 7426     Client# J.H Keeso 34 Key Chain Tally
        aryRomSerial.ary(1, 369) = "7600"

        aryRomSerial.ary(0, 370) = "50F0063006B00000014258D8137" ' 99EX Serial# 14258D8137   HH# 7444     Client# AHI Augusta end tally
        aryRomSerial.ary(1, 370) = "7600"



        aryRomSerial.ary(0, 371) = "50F0063006B00000014316D8245" ' 99EX Serial# 14316D8245   HH# 7456     Client# er Lumber - 99EX 55Key
        aryRomSerial.ary(1, 371) = "7600"

        aryRomSerial.ary(0, 371) = "50F0063006B00000014258D8041" ' 99EX Serial# 14316D8245   HH# 7456     Client# er Lumber - 99EX 55Key
        aryRomSerial.ary(1, 371) = "7600"
        
        
        aryRomSerial.ary(0, 372) = "50F0063006B00000014258D8041" ' 99EX Serial# 14258D8041   HH# 7413     Client# Gilkey Lumber - 99EX 34Key
        aryRomSerial.ary(1, 372) = "7600"

        aryRomSerial.ary(0, 373) = "50F0063006B00000014345D832B" ' 99EX Serial# 14345D832B   HH# 7480     Client# Northwest Fitzgerald - 99EX 34Key
        aryRomSerial.ary(1, 373) = "7600"
        
        aryRomSerial.ary(0, 374) = "50F0063006B00000014345D8285" ' 99EX Serial# 14345D8285   HH# 7481     Client# Northwest Fitzgerald - 99EX 34Key
        aryRomSerial.ary(1, 374) = "7600"
        
        aryRomSerial.ary(0, 375) = "50F0063006B00000014345D85EF" ' 99EX Serial# 14345D85EF   HH# 7479     Client# Northwest Fitzgerald - 99EX 34Key
        aryRomSerial.ary(1, 375) = "7600"
        
        aryRomSerial.ary(0, 376) = "50F0063006B00000014345D83C9" ' 99EX Serial# 14345D83C9   HH# 7495     Client# Baillie Clendenin - 99EX 34Key
        aryRomSerial.ary(1, 376) = "7600"
        
        aryRomSerial.ary(0, 377) = "50F0063006B00000014345D83DF" ' 99EX Serial# 14345D83DF   HH# 7496     Client# Baillie Clendenin - 99EX 34Key
        aryRomSerial.ary(1, 377) = "7600"
        
        aryRomSerial.ary(0, 378) = "50F0063006B00000014345D8334" ' 99EX Serial# 14345D8334   HH# 7499     Client# Baillie Clendenin - 99EX 34Key
        aryRomSerial.ary(1, 378) = "7600"
        
        aryRomSerial.ary(0, 379) = "50F0063006B00000014345D83D3" ' 99EX Serial# 14345D83D3   HH# 7501     Client# Baillie Clendenin - 99EX 34Key
        aryRomSerial.ary(1, 379) = "7600"
        
        aryRomSerial.ary(0, 380) = "50F0063006B00000015017D85A1" ' 99EX Serial# 15017D85A1   HH# 7505     Client# Crownover Lumber - 99EX 55Key
        aryRomSerial.ary(1, 380) = "7600"
        
        aryRomSerial.ary(0, 381) = "50F0063006B00000015008D857D" ' 99EX Serial# 15008D857D   HH# 7473     Client# Maley and Wertz - 99EX 34Key
        aryRomSerial.ary(1, 381) = "7600"
        
        aryRomSerial.ary(0, 382) = "50F0063006B00000015008D8574" ' 99EX Serial# 15008D8574   HH# 7472     Client# Maley and Wertz - 99EX 34Key
        aryRomSerial.ary(1, 382) = "7600"
        
        aryRomSerial.ary(0, 383) = "50F0063006B00000014258D8038" ' 'HH Serial 14258D8038 HH7426 34 Key - Mark Depp - Independent Grader 4/7/15
        aryRomSerial.ary(1, 383) = "7600"
        
        aryRomSerial.ary(0, 384) = "50F0063006B00000014341D82C1" ' 'HH Serial 14341D82C1 HH7556 34 Key - AHI North Garden 4/30/15
        aryRomSerial.ary(1, 384) = "7600"
        
        aryRomSerial.ary(0, 384) = "50F0063006B00000013150D80D7" ' 'HH Serial 13150D80D7 HH7330 34 Key - CLC Hardwoods 5/11/15
        aryRomSerial.ary(1, 384) = "7600"
    
        aryRomSerial.ary(0, 385) = "50F0063006B00000014341D815D" ' 'HH Serial 14341D815D HH7564 34 Key - Maley and Wertz 5/12/15
        aryRomSerial.ary(1, 385) = "7600"

        aryRomSerial.ary(0, 386) = "50F0063006B00000014341D8263" ' 'HH Serial 14341D8263 HH7565 34 Key - Maley and Wertz 5/12/15
        aryRomSerial.ary(1, 386) = "7600"
        
        aryRomSerial.ary(0, 387) = "50F0063006B00000014341D824C" ' 'HH Serial 14341D824C HH7566 34 Key - Maley and Wertz 5/12/15
        aryRomSerial.ary(1, 387) = "7600"

        aryRomSerial.ary(0, 388) = "50F0063006B00000014342D8027" ' 'HH Serial 14342D8027 HH7567 34 Key - Maley and Wertz 5/12/15
        aryRomSerial.ary(1, 388) = "7600"
        
        aryRomSerial.ary(0, 389) = "50F0063006B00000014255D84CA" ' 'HH Serial 14255D84CA HH7575 34 Key - Atco 6/15/15
        aryRomSerial.ary(1, 389) = "7600"
        
        aryRomSerial.ary(0, 390) = "50F0063006B00000014254D84AB" ' 'HH Serial 14254D84AB HH7574 34 Key - Atco 6/15/15
        aryRomSerial.ary(1, 390) = "7600"

        aryRomSerial.ary(0, 391) = "50F0063006B00000014281D82C0" ' 'HH Serial 14281D82C0 HH7628 34 Key - Kirkham/Maley Wirtz 7/30/15
        aryRomSerial.ary(1, 391) = "7600"

        aryRomSerial.ary(0, 392) = "50F0063006B00000014281D8315" ' 'HH Serial 14281D8315 HH7629 34 Key - Kirkham/Maley Wirtz 7/30/15
        aryRomSerial.ary(1, 392) = "7600"

        aryRomSerial.ary(0, 393) = "14258D80EA" ' 'Dell Axim Palm Pilot messing around
        aryRomSerial.ary(1, 393) = "7600"

        aryRomSerial.ary(0, 392) = "50F0063006B00000014281D82B3" ' 'HH Serial 14281D82B3 HH7627 34 Key - Cummings Flooring for Estimated Flooring Bundles 8/25/15 - PCC for Nic
        aryRomSerial.ary(1, 392) = "7600"
        
        aryRomSerial.ary(0, 392) = "50F0063006B00000014281D832A" ' 'HH Serial 14281D832A HH7634 34 Key - Loggers Inc. 9/30/15 - Nic
        aryRomSerial.ary(1, 392) = "7600"
        
        aryRomSerial.ary(0, 393) = "50F0063006B00000015008D837A" ' 'HH Serial 15008D837A HHL7471 34 Key - Deveraux 10/16/15 - Nic
        aryRomSerial.ary(1, 393) = "7600"

        aryRomSerial.ary(0, 394) = "50F0063006B00000015008D837E" ' 'HH Serial 15008D837E HHL7476 34 Key - John Boos 10/22/15 - Nic
        aryRomSerial.ary(1, 394) = "7600"
    
        aryRomSerial.ary(0, 395) = "50F0063006B00000015169D867C" ' 'HH Serial 15169D867C HH7657 55 Key - Gift Lumber 10/29/15 - Nic
        aryRomSerial.ary(1, 395) = "7600"
    
        aryRomSerial.ary(0, 396) = "50F0063006B00000013113D8253" 'HH Serial  13113D8253 HH2316 55 Key - Baillie Titusville (This is our 55key loaner) Lumber 11/13/15 -
        aryRomSerial.ary(1, 396) = "7600"
        
        aryRomSerial.ary(0, 397) = "50F0063006B00000014359D8292" 'HH Serial  14359D8292 HH7542 55 Key - Green Ridge 12/03/15 - Nic
        aryRomSerial.ary(1, 397) = "7600"
        
        aryRomSerial.ary(0, 398) = "50F0063006B00000014362D80F6" 'HH Serial  14362D80F6 HH7671 55 Key - Green Ridge new HH 1/11/16 - Nic
        aryRomSerial.ary(1, 398) = "7600"
        
        aryRomSerial.ary(0, 399) = "50F0063006B00000014362D80EC" 'HH Serial  14362D80EC HH7670 34 Key - T&G Lumber 1/13/16 - Nic
        aryRomSerial.ary(1, 399) = "7600"
        
        aryRomSerial.ary(0, 400) = "50F0063006B00000014281D82F0" 'HH Serial  14281D82F0 HH7694 34 Key - Baillie St. James 2/1/16 - Nic
        aryRomSerial.ary(1, 400) = "7600"
        
        aryRomSerial.ary(0, 401) = "50F0063006B00000014341D8153" 'HH Serial  14341D8153 HH7693 34 Key - Townsend Lumber 2/8/16 - Nic
        aryRomSerial.ary(1, 401) = "7600"
        
        aryRomSerial.ary(0, 402) = "50F0063006B00000014281D82CC" 'HH Serial  14281D82CC HH7692 34 Key - Lewis Lumber Products 2/8/16 - Nic
        aryRomSerial.ary(1, 402) = "7600"
        
        aryRomSerial.ary(0, 403) = "50F0063006B00000015147D8389" 'HH Serial  15147D8389 HH7702 34 Key - Green Ridge Forest Products 2/18/16 - Rob
        aryRomSerial.ary(1, 403) = "7600"
        
        aryRomSerial.ary(0, 404) = "50F0063006B00000014359D8292" 'HH Serial  14359D8292 HH7542 34 Key - Eagle Hardwoods Inc 3/08/16 - Rob
        aryRomSerial.ary(1, 404) = "7600"
        
        aryRomSerial.ary(0, 405) = "0306777" 'type: 9500L0P 43key desc: Just for a loaner S/N: 0306777
        aryRomSerial.ary(1, 405) = "9500"
        
        aryRomSerial.ary(0, 406) = "50F0063006B00000016089D8241" 'HH Serial  16089D8241 HH7711 34 Key - Neador Wood Products 4/12/16 - Nic
        aryRomSerial.ary(1, 406) = "7600"
        '
        aryRomSerial.ary(0, 407) = "50F0063006B00000014171D8238" 'HH Serial  14171D8238 HH7397 Horizon - Didn't get Licensed at some point  CRM Says 34 Key - Horizon Wood Products delivered 8/5/14 based upon crm records and licensed on 4/30/16 by PCC
        aryRomSerial.ary(1, 407) = "7600"
        
        aryRomSerial.ary(0, 408) = "50F0063006B00000014105D8025" 'HH Serial  14105D8025 HH7377 34 Key - Kamps Pure Hardwoods - delivered 04/08/2016 based upon crm records and licensed on 5/13/16 by RTM
        aryRomSerial.ary(1, 408) = "7600"
        
        aryRomSerial.ary(0, 409) = "50F0063006B00000016089D8257" 'HH Serial  16089D8257 HH7733 34 Key - Kamps Pure Hardwoods - licensed on 5/13/16 by RTM
        aryRomSerial.ary(1, 409) = "7600"
        
        aryRomSerial.ary(0, 410) = "50F0063006B00000016113D80E3" 'HH Serial  16113D80E3 HH7730 55 Key - Superior - licensed on 5/13/16 by RTM
        aryRomSerial.ary(1, 410) = "7600"

        aryRomSerial.ary(0, 411) = "50F0063006B00000014286D80B1" 'HH Serial  14286D80B1 HH7741 55 Key - Endeavor - licensed on 5/26/16 by Nic
        aryRomSerial.ary(1, 411) = "7600"
        
        aryRomSerial.ary(0, 412) = "50F0063006B00000016149D806C" 'HH Serial  16149D806C HH7750 34 Key - S&J Lumber - licensed on 6/15/16 by Nic
        aryRomSerial.ary(1, 412) = "7600"
        
        aryRomSerial.ary(0, 413) = "50F0063006B00000016149D809D" 'HH Serial  16149D809D HH7753 34 Key - Valley Hardwoods - licensed on 6/28/16 by Nic
        aryRomSerial.ary(1, 413) = "7600"
                
        aryRomSerial.ary(0, 414) = "50F0063006B00000016140D81ED" 'HH Serial  16140D81ED HH7744 34 Key - Kendrick Forest Products - licensed on 6/28/16 by Ben
        aryRomSerial.ary(1, 414) = "7600"
                
        aryRomSerial.ary(0, 415) = "50F0063006B00000016149D8090" 'HH Serial  16149D8090 HH7749 34 Key - Kendrick Forest Products - licensed on 6/28/16 by Ben
        aryRomSerial.ary(1, 415) = "7600"
        
        aryRomSerial.ary(0, 416) = "50F0063006B00000016149D806A" 'HH Serial  16149D806A HH7751 34 Key - Kendrick Forest Products - licensed on 6/28/16 by Ben
        aryRomSerial.ary(1, 416) = "7600"
                
        aryRomSerial.ary(0, 417) = "50F0063006B00000016149D80A3" 'HH Serial  16149D80A3 HH7752 34 Key - Kendrick Forest Products - licensed on 6/28/16 by Ben
        aryRomSerial.ary(1, 417) = "7600"
        
        aryRomSerial.ary(0, 418) = "50F0063006B00000016149D8081" 'HH Serial  16149D8081 HH7755 34 Key - Kendrick Forest Products - licensed on 6/28/16 by Ben
        aryRomSerial.ary(1, 418) = "7600"
                
        aryRomSerial.ary(0, 419) = "50F0063006B00000016149D8080" 'HH Serial  16149D8080 HH7756 34 Key - Kendrick Forest Products - licensed on 6/28/16 by Ben
        aryRomSerial.ary(1, 419) = "7600"
        
        aryRomSerial.ary(0, 420) = "50F0063006B00000016091D81CF" 'HH Serial  16091D81CF HH7726 55 Key - S&J Lumber - licensed on 7/7/16 by Nic
        aryRomSerial.ary(1, 420) = "7600"
        
        aryRomSerial.ary(0, 421) = "50F0063006B00000016140D81EA" 'HH Serial  16140D81EA HH7745 34 Key - Deer Park Lumber - licensed on 7/7/16 by Nic
        aryRomSerial.ary(1, 421) = "7600"
        
        aryRomSerial.ary(0, 422) = "50F0063006B00000016208D80E8" 'HH Serial  16208D80E8 HH7830 34 Key - Gates (ISIS Integration) Custom Milling - Lumber/Demo Pre Implement Licensed 9/19/16 PCC/Nic
        aryRomSerial.ary(1, 422) = "7600"
        
        aryRomSerial.ary(0, 423) = "50F0063006B00000016089D8245" 'HH Serial 16089D8245 HHD7737 Demo Unit for Mike Johnson
        aryRomSerial.ary(1, 423) = "7600"
        
        aryRomSerial.ary(0, 424) = "50F0063006B00000016291D80DC" 'HH Serial 16291D80DC HH7894 Demo Unit for Mike Johnson
        aryRomSerial.ary(1, 424) = "7600"
        
        aryRomSerial.ary(0, 425) = "50F0063006B00000016178D809A" 'HH Serial 16178D809A HH7896 34-Key eLIMBS Ink Unit for Baillie Smyrna - RM 11/15/16
        aryRomSerial.ary(1, 425) = "7600"
        
        aryRomSerial.ary(0, 426) = "50F0063006B00000016291D80FA" 'HH Serial 16291D80FA HH7912 55-Key eLIMBS Ink Unit for Baillie Clendenin - RM 12/05/16
        aryRomSerial.ary(1, 426) = "7600"

        aryRomSerial.ary(0, 427) = "50F0063006B00000016258D81A4" 'HH Serial 16258D81A4 HH7914 55-Key - Hillcrest Lumber LTD - RM 12/08/16
        aryRomSerial.ary(1, 427) = "7600"
        
        aryRomSerial.ary(0, 428) = "50F0063006B00000016258D81C6" 'HH Serial 16258D81C6 HH7917 34 Key - Gates Custom Milling 12/13/2016
        aryRomSerial.ary(1, 428) = "7600"
        
        aryRomSerial.ary(0, 429) = "50F0063006B00000016280D8048" 'HH Serial 16280D8048 HH7923 55 Key - Superior Hardwoods 12/19/2016
        aryRomSerial.ary(1, 429) = "7600"
        
        
        '**************************   2017   ******************************************** Below Here
        'Lietchfield 99EX - Nic Installing onsite 1/10/17
        aryRomSerial.ary(0, 430) = "50F0063006B00000016149D8092" 'HH Serial 16149D8092 HH7925 34 Key - Lietchfield 1/10/17
        aryRomSerial.ary(1, 430) = "7600"
        
        aryRomSerial.ary(0, 431) = "50F0063006B00000016198D80E9" 'HH Serial 16198D80E9 HH7847 34 Key - Stella-Jones Fulton, KY (Chet) BJP 1/19/2017
        aryRomSerial.ary(1, 431) = "7600"
        
        aryRomSerial.ary(0, 432) = "50F0063006B00000016178D8057" 'HH Serial 16178D8057 HH7901 34 Key - Baillie World Wood  NAJ 1/30/2017
        aryRomSerial.ary(1, 432) = "7600"
        
        aryRomSerial.ary(0, 433) = "50F0063006B00000016201D80AF" 'HH Serial 16201D80AF HH7851 34 Key - Stella-Jones Olivia  NAJ 2/1/2017
        aryRomSerial.ary(1, 433) = "7600"
        
        aryRomSerial.ary(0, 434) = "50F0063006B00000016178D8095" 'HH Serial 16178D8095 HH7852 34 Key - Stella-Jones   NAJ 2/17/2017
        aryRomSerial.ary(1, 434) = "7600"
        
        aryRomSerial.ary(0, 435) = "50F0063006B00000016140D8151" 'HH Serial 16140D8151 HH7864 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 435) = "7600"

        aryRomSerial.ary(0, 436) = "50F0063006B00000016198D80E8" 'HH Serial 16198D80E8 HH7863 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 436) = "7600"

        aryRomSerial.ary(0, 437) = "50F0063006B00000016140D8158" 'HH Serial 16140D8158 HH7862 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 437) = "7600"

        aryRomSerial.ary(0, 438) = "50F0063006B00000016178D806C" 'HH Serial 16178D806C HH7861 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 438) = "7600"

        aryRomSerial.ary(0, 439) = "50F0063006B00000016198D8170" 'HH Serial 16198D8170 HH7860 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 439) = "7600"

        aryRomSerial.ary(0, 440) = "50F0063006B00000016198D80F1" 'HH Serial 16198D80F1 HH7859 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 440) = "7600"

        aryRomSerial.ary(0, 441) = "50F0063006B00000016140D81F3" 'HH Serial 16140D81F3 HH7858 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 441) = "7600"

        aryRomSerial.ary(0, 442) = "50F0063006B00000016198D80DF" 'HH Serial 16198D80DF HH7857 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 442) = "7600"

        aryRomSerial.ary(0, 443) = "50F0063006B00000016198D80DD" 'HH Serial 16198D80DD HH7856 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 443) = "7600"

       
        aryRomSerial.ary(0, 444) = "50F0063006B00000016140D814C" 'HH Serial 16140D814C HH7855 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 444) = "7600"

        aryRomSerial.ary(0, 445) = "50F0063006B00000016207D8106" 'HH Serial 16207D8106 HH7854 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 445) = "7600"

        aryRomSerial.ary(0, 446) = "50F0063006B00000016178D8053" 'HH Serial 16178D8053 HH7853 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 446) = "7600"

        aryRomSerial.ary(0, 447) = "50F0063006B00000016198D80F8" 'HH Serial 16198D80F8 HH7850 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 447) = "7600"

        aryRomSerial.ary(0, 448) = "50F0063006B00000016198D80FE" 'HH Serial 16198D80FE HH7849 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 448) = "7600"

        aryRomSerial.ary(0, 449) = "50F0063006B00000016198D80E2" 'HH Serial 16198D80E2 HH7848 34 Key - Stella-Jones   NAJ 3/7/2017
        aryRomSerial.ary(1, 449) = "7600"
         
        aryRomSerial.ary(0, 450) = "50F0063006B00000016178D8084" 'HH Serial 16178D8084 HH7949 34 Key - Emporium Hardwoods
        aryRomSerial.ary(1, 450) = "7600"


        aryRomSerial.ary(0, 451) = "50F0063006B00000016178D8115" 'HH Serial 16178D8115 HH7937 34 Key - Prime Lumber  3/27/17
        aryRomSerial.ary(1, 451) = "7600"
        

        aryRomSerial.ary(0, 452) = "50F0063006B00000016178D80E4" 'HH Serial 16178D80E4 HH7939 34 Key - Prime Lumber 3/27/17
        aryRomSerial.ary(1, 452) = "7600"
        
        aryRomSerial.ary(0, 453) = "50F0063006B00000016178D803E" 'HH Serial 16178D803E HH7952 34 Key - Prime Lumber 4/4/17 -NAJ
        aryRomSerial.ary(1, 453) = "7600"
        
        aryRomSerial.ary(0, 454) = "50F0063006B00000016258D8171" 'HH Serial 16258D8171 HH7962 34 Key - Prime Lumber 5/1/17
        aryRomSerial.ary(1, 454) = "7600"
       
        aryRomSerial.ary(0, 455) = "50F0063006B00000016207D80D6" 'HH Serial 16207D80D6 HHD7802 - Ben's demo handheld being left at 34 Key - Prime Lumber 5/5/17
        aryRomSerial.ary(1, 455) = "7600"
            'above handheld to be returned/replaced by new unit when it arrives.
        
        aryRomSerial.ary(0, 456) = "50F0063006B00000016178D80CA" 'HH Serial 16178D80CA HH7968 34 Key - Prime Lumber 5/8/17
        aryRomSerial.ary(1, 456) = "7600"
        
        aryRomSerial.ary(0, 457) = "50F0063006B00000016178D8065" 'HH Serial 16178D8065 HH7977 34 Key - Matson Lumber 6/8/17 - NAJ
        aryRomSerial.ary(1, 457) = "7600"
        
        aryRomSerial.ary(0, 458) = "50F0063006B00000016178D8069" 'HH Serial 16178D8069 HH7988 34 Key - Baillie World Wood 6/29/17 (licensed on 7/12/17) - RTM
        aryRomSerial.ary(1, 458) = "7600"

        aryRomSerial.ary(0, 459) = "50F0063006B00000016178D80B9" 'HH Serial 16178D80B9 HH7990 34 Key - Pennyrile Sawmill 7/17/17 NAJ
        aryRomSerial.ary(1, 459) = "7600"

        aryRomSerial.ary(0, 460) = "50F0063006B00000016178D8067" 'HH Serial 16178D8067 HH8002 34 Key - AHI - Graham, TN Sawmill 8/26/17 NAJ
        aryRomSerial.ary(1, 460) = "7600"
        
        aryRomSerial.ary(0, 461) = "50F0063006B00000016178D808B" 'HH Serial 16178D8088 HH8004 34 Key - AHI Graham,TN Sawmill 8/26/17 NAJ
        aryRomSerial.ary(1, 461) = "7600"
        
        'September 2017 Below Here ****
        
        aryRomSerial.ary(0, 462) = "50F0063006B00000016178D8052"  'HH Serial 16178D8052 HH7986 34 Key Stoney Point 6/20/2017 - RDW
        aryRomSerial.ary(1, 462) = "7600"
    
        aryRomSerial.ary(0, 463) = "50F0063006B00000016178D8083"  'HH Serial 16178D8083 HH7996 34 Key Cobble Creek 7/25/2017 - RDW
        aryRomSerial.ary(1, 463) = "7600"
    
        aryRomSerial.ary(0, 464) = "50F0063006B00000017320D82D3"  'HH Serial 17320D82D3 HH8171 34 Key Baillie Leitchfield 12/15/2017 - RDW
        aryRomSerial.ary(1, 464) = "7600"
    
        aryRomSerial.ary(0, 465) = "50F0063006B00000017320D8288"  'HH Serial 17320D8288 HH8170 34 Key Baillie Leitchfield 12/15/2017 - RDW
        aryRomSerial.ary(1, 465) = "7600"
    
        aryRomSerial.ary(0, 466) = "50F0063006B00000016178D8069"  'HH Serial 16178D8069 HH7988 34 Key Baillie World Wood 6/29/2017 - RDW
        aryRomSerial.ary(1, 466) = "7600"
    
        aryRomSerial.ary(0, 467) = "50F0063006B00000016178D80CD"  'HH Serial 16178D80CD HH8010 34 Key Matson Lumber Company 10/26/2017 - RDW
        aryRomSerial.ary(1, 467) = "7600"
    
        aryRomSerial.ary(0, 468) = "50F0063006B00000016178D8065"  'HH Serial 16178D8065 HH7977 34 Key Matson Lumber Company 6/7/2017 - RDW
        aryRomSerial.ary(1, 468) = "7600"
    
        aryRomSerial.ary(0, 469) = "50F0063006B00000016178D8062"  'HH Serial 16178D8062 HH7994 34 Key HAVCO Wood Products 8/3/2017 - RDW
        aryRomSerial.ary(1, 469) = "7600"

        aryRomSerial.ary(0, 470) = "50F0063006B00000017293D8005" 'HH Serial 17293D8005 HH8059 34 Key - Gates Custom Milling 11/30/2017 - RDW
        aryRomSerial.ary(1, 470) = "7600"

        aryRomSerial.ary(0, 471) = "50F0063006B00000016178D806D" 'HH Serial 16178D806D HH8012 34 Key - S&J Lumber 11/30/2017 - RDW
        aryRomSerial.ary(1, 471) = "7600"

        aryRomSerial.ary(0, 472) = "50F0063006B00000017327D80BD" 'HH Serial 17327D80BD HH8186 55 Key - S&J Lumber 1/31/2018 - RDW
        aryRomSerial.ary(1, 472) = "7600"

        aryRomSerial.ary(0, 473) = "50F0063006B00000017293D8016" 'HH Serial 17293D8016 HH8176 34 Key - Lincoln County Hardwoods 2/5/2018 - RDW
        aryRomSerial.ary(1, 473) = "7600"

        aryRomSerial.ary(0, 474) = "50F0063006B00000017293D801E" 'HH Serial 17293D801E HH8179 34 Key - Shetler Lumber 2/5/2018 - RDW
        aryRomSerial.ary(1, 474) = "7600"

        aryRomSerial.ary(0, 475) = "50F0063006B00000017293D8006" 'HH Serial 17293D8006 HH8180 34 Key - Hardy Valley Lumber 2/19/2018 - RDW
        aryRomSerial.ary(1, 475) = "7600"

        aryRomSerial.ary(0, 476) = "50F0063006B00000016178D80BD" 'HH Serial 16178D80BD HH8011 34 Key - Krueger Lumber Company 10/27/2017 - RDW
        aryRomSerial.ary(1, 476) = "7600"

        aryRomSerial.ary(0, 477) = "50F0063006B00000018165D8313 " 'HH Serial 18165d8313 HH8265 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 477) = "7600"
        
        aryRomSerial.ary(0, 478) = "50F0063006B00000018135D804F " 'HH Serial 18135D804F HH8266 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 478) = "7600"
        
        aryRomSerial.ary(0, 479) = "50F0063006B00000018135D807C " 'HH Serial 18135D807C HH8267 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 479) = "7600"
        
        aryRomSerial.ary(0, 480) = "50F0063006B00000018165D8198 " 'HH Serial 18165D8198 HH8268 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 480) = "7600"
        
        aryRomSerial.ary(0, 481) = "50F0063006B00000018135D8032 " 'HH Serial 18135D8032 HH8269 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 481) = "7600"
        
        aryRomSerial.ary(0, 482) = "50F0063006B00000018135D8066 " 'HH Serial 18135D8066 HH8270 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 482) = "7600"
        
        aryRomSerial.ary(0, 483) = "50F0063006B00000018135D8017 " 'HH Serial 18135D8017 HH8271 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 483) = "7600"
        
        aryRomSerial.ary(0, 484) = "50F0063006B00000018135D8081 " 'HH Serial 18135D8081 HH8272 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 484) = "7600"
        
        aryRomSerial.ary(0, 485) = "50F0063006B00000018132D80C0 " 'HH Serial 18132D80C0 HH8273 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 485) = "7600"
        
        aryRomSerial.ary(0, 486) = "50F0063006B00000018135D805A " 'HH Serial 18135D805A HH8274 34 Key - North Coutry Lumber 8/22/2018 PCC
        aryRomSerial.ary(1, 486) = "7600"
        
        aryRomSerial.ary(0, 487) = "50F0063006B00000018165D819B " 'HH Serial 18165D819B HH8363 34 Key - Appalachian Hardwood 1/23/2019 NAJ
        aryRomSerial.ary(1, 487) = "7600"
        
    CheckForHandheldLicense = True
    
    Exit Function
ErrorHandler:
    CheckForHandheldLicense = False
    MsgBox "Error in CheckForHandheldLicense: " & Err.Number & "-" & Err.Description
    Exit Function
    
End Function
Public Function ValidateAID_Any(strSearchValue As String, strTypeToValidate As String, strSearchType As String, _
                                strReturnField As String, _
                                Optional lngReturnID As Long, Optional strReturnAID As String, _
                                Optional strReturnName As String) As String

On Error GoTo ErrorHandler

    lngReturnID = 0
    strReturnAID = ""
    strReturnName = ""
    
    If SC(strSearchValue, "") = True Then
        lngReturnID = 0
        strReturnAID = ""
        strReturnName = ""
        
    Else
            
        Select Case tcu(strTypeToValidate)
            
            Case Is = "VENDOR"
            
                ValidateAID_Any = GetOrgData(strSearchValue, strSearchType, strTypeToValidate, strReturnField, lngReturnID, strReturnAID, strReturnName)
                Exit Function
            Case Is = "CARRIER"
                If SC(gSettings.Org_IncludeAllVendorsInCarrierList, "YES") = True Then
                    strTypeToValidate = "VENDOR"
                Else
                    'leave it as carrier/as it was sent in
                End If
                ValidateAID_Any = GetOrgData(strSearchValue, strSearchType, strTypeToValidate, strReturnField, lngReturnID, strReturnAID, strReturnName)
                
            Case Is = "PRODID"
            
            
            Case Is = "SPECIES"
                Call GetSpeciesData(strSearchValue, strSearchType, strReturnField, strReturnAID, strReturnName, lngReturnID)
                
            Case Is = "GRADE"
                Call GetGradeData(strSearchValue, strSearchType, strReturnField, lngReturnID, strReturnAID, strReturnName)
            Case Is = "THK", "THICKNESS"
                Call GetThicknessData(strSearchValue, strSearchType, strReturnField, lngReturnID, strReturnAID, strReturnName)
            Case Is = "ORG"
                ValidateAID_Any = GetOrgData(strSearchValue, strSearchType, strTypeToValidate, strReturnField, lngReturnID, strReturnAID, strReturnName)
                Exit Function
            Case Is = "STATUS"
                ValidateAID_Any = GetStatusData(strSearchValue, strSearchType, strReturnField)
                Exit Function
            Case Else
                Exit Function
        End Select
    End If
    
    
    
    'Return requested field to calling module whether valid entry found or not.
    'if sub not exited above, then exit and return requested value
    If SC(strReturnField, "ID") = True Then
        ValidateAID_Any = CStr(lngReturnID)
    ElseIf SC(strReturnField, "AID") = True Or SC(strReturnField, "HHAID") = True Then
        ValidateAID_Any = strReturnAID
    ElseIf SC(strReturnField, "NAME") Or SC(strReturnField, "DESC") = True Then
        ValidateAID_Any = strReturnName
    Else
        ValidateAID_Any = strReturnAID
    End If
    
    
    Exit Function
ErrorHandler:
    If SC(strReturnField, "ID") = True Then
        ValidateAID_Any = "-1"
    Else
        ValidateAID_Any = "NOTFOUND"
    End If
    
    ValidateAID_Any = "-1"
    MsgBox "Error in ValidateAID_Any SearchFor=" & strSearchValue & "  -  Search Type=" & strTypeToValidate & vbCrLf & Err.Number & "-" & Err.Description
    Exit Function
    
End Function


Public Sub HHMODEL_KEYCODE_ASSIGNMENT(strModel As String, strKeypad As String)

On Error GoTo ErrorHandler

    '**********SET Default values from the original Constants SEtup for the keys below, then override whichever shouldn't be active/set/conflict
    '********** DO NOT OVERRIDE THE STANDARD F1,F2,F3 etc..or the keyboard won't work when you type on pc keyboard, or in debug/testing/code modes.
    
    KeyHelp2_99EX = 112
    KeyDelete = 189     'HHP Blue/Del (-)
    ''Find default values for above to and set them
    ''***********************
    
    keyEnter = 13      'ENTER ONLY
    KeyUp = 38        'NOTHING
    KeyFldExit = 149
    
    KeyDown = 40      'NOTHING
    keyLeft = 37
    keyRight = 39
    
    keyAD = 65         'HHP A
    keyKD = 74         'HHP J
    keyCancel = 27     'HHP ESC Key
    KeyHelp = 112      'HHP F1
    KeyHelp_9900 = 227 'HHP 9900 F1 Key
    KeyHelp2 = 45      'HHP INS Key
    
    keySave = 113      'HHP F2
    
    KeyView = 114      'HHP F3
    
    KeyEdit = 115      'HHP F4
    
    
    keyAdd = 187    'HHP Blue/SP (+)
    
    KeyMax = 200   'Place Holder for now
    KeyMin = 200
    
    KeyF1 = 112
    KeyF2 = 113
    KeyF3 = 114
    KeyF4 = 115
    KeyF5 = 116
    KeyF6 = 96
    KeyF7 = 97
    KeyF8 = 119
    KeyF9 = 120
    KeyF10 = 121
    
    KeyF1_9900 = 227
    KeyF2_9900 = 228
    KeyF3_9900 = 230
    KeyF4_9900 = 233
    KeyF5_9900 = 234
    KeyF6_9900 = 235
    KeyF7_9900 = 236
    KeyF8_9900 = 237
    KeyF9_9900 = 238
    KeyF10_9900 = 239
    KeyScan = 42
    
    KeySP = 32
    KeyBKSP = 8
    KeyCTRL = 17
    KeySymbolRedDot = 126 'This is actual equivalent to the "~"
    KeyTab = 9 'Tab key on the HHP Handhelds
    KeyComma = 188 'This is true on 56 key, next to zero button
    KeyDecimal = 190  'This is true on 56 key, next to zero button
    KeySemicolon_99EX = 186
    KeyPoundSign_99EX = 155
    '**********************************END OF DEFAULT SETTINGS
    
    If SC(strModel, "99EX") = True Then
        'only override the keys that are necessary to override
        'currently these are same as defaults ,but these likely need changed..
            'especially keyad/keykd/keyview (with alpha on bundle today , 2 button doesn't do what "2" does with alpha off)
        
        keyAD = 65         'HHP A
        keyKD = 74         'HHP J
        KeyHelp = 112      'HHP F1
        KeyHelp_9900 = 227 'HHP 9900 F1 Key
        KeyHelp2 = 45      'HHP INS Key
        keySave = 113      'HHP F2
        KeyView = 114      'HHP F3
        KeyEdit = 115      'HHP F4
        keyAdd = 187    'HHP Blue/SP (+)
        
        KeyDelete = 46
        If SC(gSettings.KeySpaceAsF1, "YES") Then
            KeyHelp2_99EX = 32
        ElseIf IsNumeric(gSettings.KeySpaceAsF1) = True Then
            KeyHelp2_99EX = CInt(gSettings.KeySpaceAsF1)
        End If
    Else
    
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error in HHModel_Keycode_Assignment " & Err.Number & "-" & Err.Description
    Exit Sub
    
End Sub

Public Function ShowReport_PrintGridForm(frmReturn As Form, strReportType As String, strReportTitle As String, _
                                         Optional strReportFileName As String, Optional strVariable1 As String, _
                                         Optional strVariable2 As String, Optional strVariable3 As String, _
                                         Optional strVariable4 As String) As Boolean
On Error GoTo ErrorHandler

    Set frmReport_PrintGrid.pfrmReturn = frmReturn
    frmReport_PrintGrid.pstrReportTitle = strReportTitle
    frmReport_PrintGrid.pstrPrintType = strReportType
    
    
    frmReport_PrintGrid.pstrPrintReportVariable1 = strVariable1
    frmReport_PrintGrid.pstrPrintReportVariable2 = strVariable2
    frmReport_PrintGrid.pstrPrintReportVariable3 = strVariable3
    frmReport_PrintGrid.pstrPrintReportVariable4 = strVariable4
    
    frmReport_PrintGrid.pstrPrintFileName = strReportFileName
    frmReport_PrintGrid.Show
    
    ShowReport_PrintGridForm = True
    
    Exit Function

ErrorHandler:
    ShowReport_PrintGridForm = False


End Function

