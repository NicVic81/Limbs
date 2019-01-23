VERSION 5.00
Object = "{1A028BB9-0262-48D8-9B67-D74CE4C2A7E8}#6.0#0"; "IngotTextBoxCtl.dll"
Object = "{BAE83016-CA11-4D8A-BBFA-0AE9863B82DE}#6.0#0"; "IngotLabelCtl.dll"
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#6.0#0"; "IngotButtonCtl.dll"
Object = "{54171E73-16D6-4021-A96E-E59C0FE780E9}#6.0#0"; "IngotShapeCtl.dll"
Object = "{F9885939-2FBB-491F-8EC3-DBC61CCFA7DB}#6.0#0"; "IngotGraphicCtl.dll"
Object = "{1A298ECE-60AD-4C91-BB12-3092A2250D21}#6.0#0"; "IngotTimerCtl.dll"
Begin VB.Form frmLogin 
   BackColor       =   &H00ECF7F7&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3660
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   269
   ScaleMode       =   0  'User
   ScaleWidth      =   239
   Begin IngotLabelCtl.AFLabel lblVersion 
      Height          =   285
      Left            =   2430
      OleObjectBlob   =   "frmLogin.frx":164A
      TabIndex        =   14
      Top             =   810
      Width           =   1230
   End
   Begin IngotLabelCtl.AFLabel lblLeft 
      Height          =   4470
      Left            =   0
      OleObjectBlob   =   "frmLogin.frx":1691
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin IngotLabelCtl.AFLabel lblBorder 
      Height          =   195
      Left            =   0
      OleObjectBlob   =   "frmLogin.frx":16D6
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   3660
   End
   Begin IngotTextBoxCtl.AFTextBox lblUserLoginNameHH 
      Height          =   345
      Left            =   1170
      OleObjectBlob   =   "frmLogin.frx":171B
      TabIndex        =   11
      Tag             =   "txtUserLoginNameHH"
      Top             =   1620
      Width           =   990
   End
   Begin IngotTextBoxCtl.AFTextBox txtUserLoginNameHH 
      Height          =   345
      Left            =   240
      OleObjectBlob   =   "frmLogin.frx":1778
      TabIndex        =   2
      Tag             =   "txtUserLoginNameHH"
      Top             =   1620
      Width           =   945
   End
   Begin IngotLabelCtl.AFLabel AFLabel1 
      Height          =   285
      Left            =   945
      OleObjectBlob   =   "frmLogin.frx":17D5
      TabIndex        =   10
      Top             =   1035
      Width           =   1725
   End
   Begin IngotButtonCtl.AFButton AFButton1 
      Height          =   600
      Left            =   2475
      OleObjectBlob   =   "frmLogin.frx":1823
      TabIndex        =   9
      Tag             =   "cmdLogin"
      Top             =   2115
      Width           =   1125
   End
   Begin IngotTimerCtl.AFTimer tmrBackup 
      Height          =   480
      Left            =   3060
      OleObjectBlob   =   "frmLogin.frx":186C
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3510
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotButtonCtl.AFButton cmdLogin 
      Height          =   645
      Left            =   2475
      OleObjectBlob   =   "frmLogin.frx":188D
      TabIndex        =   6
      Tag             =   "cmdLogin"
      Top             =   1350
      Width           =   1125
   End
   Begin IngotShapeCtl.AFShape AFShape4 
      Height          =   135
      Left            =   120
      OleObjectBlob   =   "frmLogin.frx":18D7
      TabIndex        =   5
      Top             =   4800
      Width           =   3390
   End
   Begin IngotTextBoxCtl.AFTextBox txtUserPasswordHH 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "frmLogin.frx":1910
      TabIndex        =   4
      Tag             =   "txtUserPasswordHH"
      Top             =   2325
      Width           =   1935
   End
   Begin IngotLabelCtl.AFLabel lblUserPasswordHH 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmLogin.frx":196E
      TabIndex        =   3
      Top             =   2085
      Width           =   1935
   End
   Begin IngotLabelCtl.AFLabel lbl 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmLogin.frx":19BB
      TabIndex        =   0
      Top             =   1350
      Width           =   1935
   End
   Begin IngotGraphicCtl.AFGraphic AFGraphic1 
      Height          =   4125
      Left            =   135
      OleObjectBlob   =   "frmLogin.frx":1A09
      TabIndex        =   1
      Top             =   225
      Width           =   5490
   End
   Begin IngotLabelCtl.AFLabel lblBackup 
      Height          =   150
      Left            =   135
      OleObjectBlob   =   "frmLogin.frx":1A4B
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   3660
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fLoad As Boolean
Option Explicit
'

Private Sub AFButton1_Click()
    End
End Sub

Private Sub cmdLogin_Click()
    
    Dim cbo As AFComboBox
    Dim intError As Integer
''
    
On Error GoTo ErrorHandler

    If Trim(LCase(txtUserLoginNameHH.text)) = "import" And Trim(LCase(txtUserPasswordHH.text)) = "import" Then

    ElseIf Trim(LCase(txtUserLoginNameHH.text)) = "key" And Trim(LCase(txtUserPasswordHH.text)) = "key" Then
        frmKeyPress.Show
        Exit Sub
    ElseIf Trim(txtUserLoginNameHH.text) = "" And Trim(txtUserPasswordHH.text) = "" Then
        MsgBox "You must enter a username and password to continue."
        txtUserLoginNameHH.SetFocus
        Exit Sub
        '
    ElseIf Trim(LCase(txtUserLoginNameHH.text)) = "xxx" And Trim(LCase(txtUserPasswordHH.text)) = "4275" Then
        End
    End If


intError = 1
    If UserPasswordCheck(Trim(LCase(txtUserLoginNameHH.text)), Trim(LCase(txtUserPasswordHH.text))) = False Then
        MsgBox "You have entered an invalid user name or password.  Please try again!"
        txtUserLoginNameHH.SetFocus
    Else
intError = 6
    
    'MsgBox "Scanner"
        If gScanTags = True And gstrScanner <> "SCANWEDGE" Then load frmScanner
        
        'Top if added for new pallet one version 8/6/12
        If sc(gSettings.TallyFormVersion, "Horiz") = True Then
            frmMainMenuHoriz.Show
        ElseIf gSettings.TallyFormVersion = "1" Or gSettings.TallyFormVersion = "2" Then
            frmMenuMain.Show
            
        ElseIf (gstrHHModel = "99EX" Or gSettings.TallyHeaderVersion = "3") Then
            frmMainMenuNew.Show
        Else
            frmMenuMain.Show
        End If
        Unload Me
    End If

    Exit Sub

ErrorHandler:
    MsgBox "intError Position: " & intError & " - " & Err.Number & "-" & Err.Description
    Exit Sub

End Sub

Private Sub Form_Activate()
Dim I As Integer
Dim tutil As ttblLogUtilityRecord
    
    If gstrHHModel = "99EX" Or gSettings.TallyHeaderVersion = "3" Then
        Me.Top = 17
        Me.Left = 0
    End If
    If fLoad = False Then Call txtUserLoginNameHH_LostFocus
    txtUserLoginNameHH.SetFocus
    
    'Be sure all the databases are closed
    closealldatabases


    txtUserLoginNameHH.SetFocus
    
    Call DBupdate20131001
    
    loadgaryTallyFields
    InitializeGlobalVariables
    
    #If APPFORGE Then
        
    #Else
        gstrHHModel = "99EX"
    #End If
    
    If gSettings.SkipLogin = "YES" Then
        txtUserLoginNameHH.text = "001"
        txtUserPasswordHH.text = "123"
        Call cmdLogin_Click
   End If
   
    

   ' If keypad43 = True Then
    
   ' txtUserLoginNameHH.text = "001"
   ' txtUserPasswordHH.text = "123"
   ' cmdLogin_Click
    
   ' End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If fLoad = True Then
        fLoad = False
        Exit Sub
    End If

    Select Case KeyCode
        Case 10
            KeyCode = 0
        Case 13, keyEnter

            If Shift = 0 Then
                Select Case Screen.ActiveControl.tag
                    Case Is = "txtUserLoginNameHH"
                        txtUserPasswordHH.SetFocus
                    Case Is = "txtUserPasswordHH"
                        KeyCode = 0
                        Call cmdLogin_Click
                        KeyCode = 0
                    Case Is = "cmdLogin"
                        Call cmdLogin_Click
                        KeyCode = 0
                End Select
                KeyCode = 0

            ElseIf Shift = 1 Then
                Select Case Screen.ActiveControl.tag
                    Case Is = "txtUserLoginNameHH"
                        'Do Nothing...just stay here
                    Case Is = "txtUserPasswordHH"
                        txtUserLoginNameHH.SetFocus
                    Case Is = "cmdLogin"
                        txtUserPasswordHH.SetFocus
                End Select
                KeyCode = 0
            End If
        Case keyCancel
'''            If Trim(UCase(txtUserPasswordHH.Text)) = "MACOUT1" Then
                End
'''            End If
        Case KeyHelp
            If Screen.ActiveControl.tag = "txtUserLoginNameHH" Then
                frmLookup.pstrType = "USER"
                Set frmLookup.txtReturn = txtUserLoginNameHH
                Set frmLookup.pfrmReturn = Me
                frmLookup.Show
            End If
        Case KeyUp, KeyDown
            If Screen.ActiveControl.tag = "txtUserLoginNameHH" Then
                txtUserPasswordHH.SetFocus
            Else
                txtUserLoginNameHH.SetFocus
            End If

        Case Else
            Exit Sub
    End Select

End Sub

Private Sub Form_Load()
    Dim aryRomSerial() As String
    Dim fRestore As Boolean
    

keypad43 = False

    Dim I As Integer
    Dim fLicense As Boolean

    Dim FileMgr As CFileManager
    Dim TempFile As CFileBinaryReadable
    Dim cbo As AFComboBox
    Dim intError As Double
    Dim fRead As Boolean

    gfRestrictNonHandheldFunctions = False
    
'''    Call LoadDoyleFootage
    
''''    ReDim aryRomSerial(0,29)
''''    aryRomSerial(0,0) = "74233DF43F10076D1"
''''    aryRomSerial(0,1) = "D4A110744D10076D1"
''''    aryRomSerial(0,3) = "7E5151F44410076D1"
''''    aryRomSerial(0,4) = "74E261E44810076D1"
''''    aryRomSerial(0,5) = "01600D740310076D1" 'ITTW01
''''    aryRomSerial(0,6) = "0BF008842110076D1" 'DEMO 7900 UNIT
''''    aryRomSerial(0,7) = "D4C110744910076D1" 'HH1018
''''    aryRomSerial(0,8) = "D6D361D44110076D1" 'HH1019
''''    aryRomSerial(0,9) = "D6D361D44110076D1" 'HH1019
''''    aryRomSerial(0,10) = "0AB001741810076D1" 'trucS9
''''    aryRomSerial(0,11) = "03100FF42110076D1" 'HH1007
''''    aryRomSerial(0,12) = "0AB001741810076D1" 'GPTRUCS9'
''''    aryRomSerial(0,13) = "0D1000143710076D1" 'GPTRUCS6
''''    aryRomSerial(0,14) = "D4F110744E10076D1" 'HH assigned to sherwood
''''    aryRomSerial(0,15) = "4EB482808210076D1" 'HH1001
''''    aryRomSerial(0,16) = "0C6001441210076D1" 'HH1002
''''    aryRomSerial(0,17) = "0DB001741710076D1" 'HH1003
''''    aryRomSerial(0,18) = "0B8000343B10076D1" 'HH1004
''''    aryRomSerial(0,19) = "0C1000143310076D1"
''''    aryRomSerial(0,20) = "0D8000343410076D1" 'HH1010
''''    aryRomSerial(0,21) = "042002A41410076D1" 'HHP Owned DEVICE
''''    aryRomSerial(0,22) = "09D00FB42C10076D1" 'HH009
''''    aryRomSerial(0,23) = "038000343010076D1" 'HH1011
''''    aryRomSerial(0,24) = "C3F663644310076D1" 'New One 0192412
''''    aryRomSerial(0,25) = "D05366244510076D1" 'New 43 key
''''    aryRomSerial(0,26) = "D8622AD44F10076D1"
''''    aryRomSerial(0,27) = "D8422AD44710076D1"
''''    aryRomSerial(0,28) = "7D0523344810076D1"
''''    aryRomSerial(0,29) = "7D0523344810076D1"
      
'    aryRomSerial(0,0) = "74233DF43F10076D1"
'    aryRomSerial(0,1) = "D4A110744D10076D1"
'    aryRomSerial(0,3) = "7E5151F44410076D1"
'    aryRomSerial(0,4) = "74E261E44810076D1"
'    aryRomSerial(0,5) = "01600D740310076D1" 'ITTW01
'    aryRomSerial(0,6) = "0BF008842110076D1" 'DEMO 7900 UNIT
'    aryRomSerial(0,7) = "D4C110744910076D1" 'HH1018
'    aryRomSerial(0,8) = "D6D361D44110076D1" 'HH1019
'    aryRomSerial(0,9) = "D6D361D44110076D1" 'HH1019
'    aryRomSerial(0,10) = "0AB001741810076D1" 'trucS9
'    aryRomSerial(0,11) = "03100FF42110076D1" 'HH1007
'    aryRomSerial(0,12) = "0AB001741810076D1" 'GPTRUCS9'
'    aryRomSerial(0,13) = "0D1000143710076D1" 'GPTRUCS6
'    aryRomSerial(0,14) = "D4F110744E10076D1" 'HH assigned to sherwood
'    aryRomSerial(0,15) = "4EB482808210076D1" 'HH1001
'    aryRomSerial(0,16) = "0C6001441210076D1" 'HH1002
'    aryRomSerial(0,17) = "0DB001741710076D1" 'HH1003
'    aryRomSerial(0,18) = "0B8000343B10076D1" 'HH1004
'    aryRomSerial(0,19) = "0C1000143310076D1"
'    aryRomSerial(0,20) = "0D8000343410076D1" 'HH1010
'    aryRomSerial(0,21) = "042002A41410076D1" 'HHP Owned DEVICE
'    aryRomSerial(0,22) = "09D00FB42C10076D1" 'HH009
'    aryRomSerial(0,23) = "038000343010076D1" 'HH1011
'    aryRomSerial(0,24) = "C3F663644310076D1" 'New One 0192412
'    aryRomSerial(0,25) = "C3B663544710076D1" '
'    aryRomSerial(0,26) = "7D6523344810076D1"
'    aryRomSerial(0,27) = "C34663544510076D1" '
'    aryRomSerial(0,28) = "0B8000343A10076D1" ' Handled HH. delete this serial once HH issold!
'    aryRomSerial(0,29) = "C31663544A10076D1" '
'    aryRomSerial(0,30) = "00600D840E10076D1" '
'    aryRomSerial(0,31) = "C31663544A10076D1" '
'    aryRomSerial(0,32) = "D8622AD44F10076D1" '
'    aryRomSerial(0,33) = "7DA523344E10076D1" '
'    aryRomSerial(0,34) = "440037003600300030000000-4137303000" ' HHP 7600 ITTW02
'    aryRomSerial(0,35) = "7D0523344810076D1" '
'    aryRomSerial(0,36) = "0DC004D41110076D1" '
'    aryRomSerial(0,37) = "7D7523344310076D1" 'HH1023
'    aryRomSerial(0,60) = "0DB00E440710076D1"
'    aryRomSerial(0,61) = "40A663344E10076D1"
'    aryRomSerial(0,62) = "19A86EE45710076D1"
'    aryRomSerial(0,63) = "8C1061445110076D1"
'    aryRomSerial(0,64) = "8C4161445510076D1"
'    aryRomSerial(0,65) = "88D651A46410076D1"
'    aryRomSerial(0,66) = "889651A46210076D1"               ' Milton Timber
'    aryRomSerial(0,66) = "889651A46010076D1"               ' 47 Lumber
'    aryRomSerial(0,67) = "880651A46010076D1"
'    aryRomSerial(0,68) = "886651A46D10076D1"          ' 0286145
'    aryRomSerial(0,69) = "3FF62A245F10076D1"
'    aryRomSerial(0,70) = "6BE353746310076D1"
'    aryRomSerial(0,71) = "889651A46210076D1"
'    aryRomSerial(0,72) = "36E266C46E10076D1"
'    aryRomSerial(0,73) = "6B0353746F10076D1"
'    aryRomSerial(0,74) = "362266E46110076D1"
'    aryRomSerial(0,75) = "347326F46A10076D1"
'    aryRomSerial(0,76) = "006E604F30137"

ReDim aryRomSerial(4, 500)
    aryRomSerial(0, 0) = "74233DF43F10076D1"
    aryRomSerial(0, 1) = "D4A110744D10076D1"
    aryRomSerial(0, 3) = "7E5151F44410076D1"
    aryRomSerial(0, 4) = "74E261E44810076D1"
    aryRomSerial(0, 5) = "01600D740310076D1" 'ITTW01
    aryRomSerial(0, 6) = "0BF008842110076D1" 'DEMO 7900 UNIT
    aryRomSerial(0, 7) = "D4C110744910076D1" 'HH1018
    aryRomSerial(0, 8) = "D6D361D44110076D1" 'HH1019
    aryRomSerial(0, 9) = "D6D361D44110076D1" 'HH1019
    aryRomSerial(0, 10) = "0AB001741810076D1" 'trucS9
    aryRomSerial(0, 11) = "03100FF42110076D1" 'HH1007
    aryRomSerial(0, 12) = "0AB001741810076D1" 'GPTRUCS9'
    aryRomSerial(0, 13) = "0D1000143710076D1" 'GPTRUCS6
    aryRomSerial(0, 14) = "D4F110744E10076D1" 'HH assigned to sherwood
    aryRomSerial(0, 15) = "4EB482808210076D1" 'HH1001
    aryRomSerial(0, 16) = "0C6001441210076D1" 'HH1002
    aryRomSerial(0, 17) = "0DB001741710076D1" 'HH1003
    aryRomSerial(0, 18) = "0B8000343B10076D1" 'HH1004
    aryRomSerial(0, 19) = "0C1000143310076D1"
    aryRomSerial(0, 20) = "0D8000343410076D1" 'HH1010
    aryRomSerial(0, 21) = "042002A41410076D1" 'HHP Owned DEVICE
    aryRomSerial(0, 22) = "09D00FB42C10076D1" 'HH009
    aryRomSerial(0, 23) = "038000343010076D1" 'HH1011
    aryRomSerial(0, 24) = "C3F663644310076D1" 'New One 0192412
    aryRomSerial(0, 25) = "C3B663544710076D1" '
    aryRomSerial(0, 26) = "7D6523344810076D1"
    aryRomSerial(0, 27) = "C34663544510076D1" '
    aryRomSerial(0, 28) = "0B8000343A10076D1" ' Handled HH. delete this serial once HH issold!
    aryRomSerial(0, 29) = "C31663544A10076D1" '
    aryRomSerial(0, 30) = "00600D840E10076D1" '
    aryRomSerial(0, 31) = "C31663544A10076D1" '
    aryRomSerial(0, 32) = "D8622AD44F10076D1" '
    aryRomSerial(0, 33) = "7DA523344E10076D1" '
    aryRomSerial(0, 34) = "440037003600300030000000-4137303000" ' HHP 7600 ITTW02
    aryRomSerial(0, 35) = "7D0523344810076D1" '
    aryRomSerial(0, 36) = "0DC004D41110076D1" '
    aryRomSerial(0, 37) = "7D7523344310076D1" 'HH1023
    aryRomSerial(0, 60) = "0DB00E440710076D1"
    aryRomSerial(0, 61) = "40A663344E10076D1"
    aryRomSerial(0, 62) = "19A86EE45710076D1"
    aryRomSerial(0, 63) = "8C1061445110076D1"
    aryRomSerial(0, 64) = "910151546310076D1"
    aryRomSerial(0, 65) = "91C151646310076D1"
    aryRomSerial(0, 66) = "8C4161445510076D1"
    aryRomSerial(0, 67) = "88D651A46410076D1"
    aryRomSerial(0, 68) = "889651A46210076D1"     ' Milton Timber
    aryRomSerial(0, 69) = "888651A46A10076D1"    ' Leer
    aryRomSerial(0, 70) = "882651B46010076D1"
    aryRomSerial(0, 71) = "3FD62A245710076D1"
    aryRomSerial(0, 72) = "889651A46010076D1"              ' 47 Lumber
    aryRomSerial(0, 73) = "3FF62A245F10076D1" '56KEY 9500LOP-432C30E  SERIAL: 0286123
    aryRomSerial(0, 74) = "889651B46610076D1" '35KEY 9500L0P-422C30E  SERIAL: 0286088
    aryRomSerial(0, 75) = "886651A46D10076D1" '43key 9500L0P-412C30E  SERIAL: 0286145
    aryRomSerial(0, 76) = "886651B46910076D1" '35key 9500L0P-422C30E  SERIAL: 0286073
    aryRomSerial(0, 77) = "886651B46710076D1" '35key 9500L0P-422C30E  SERIAL: 0286072
    aryRomSerial(0, 78) = "88C651A46D10076D1" '43key 9500L0P-412C30E  SERIAL: 0286156
    aryRomSerial(0, 79) = "88E651A46B10076D1" '35KEY 9500L0P-422C30E  SERIAL: 0286089
    aryRomSerial(0, 80) = "1002A4016245D251A800-0050BF7A60E2" 'Atlanta Hardwoods Bruce manalan Symbol Device #1
    aryRomSerial(0, 81) = "077192003046B9B1D800-0050BF7A60E2" 'Atlanta Hardwoods Part #  MC9090-GJ0HBFGA2WR   S/N 8045000502526 3/18/08
    aryRomSerial(0, 82) = "0392A9010347CB911800-0050BF7A60E2" 'Atlanta Hardwoods Part # MC9090-GJ0HBFGA2WR   s/n 8010000505870   3/18/08
    aryRomSerial(0, 83) = "6B0353746910076D1" '35KEY 9500L0P-422C30E Serial: 0306570  Arrived 3/18/08
    aryRomSerial(0, 84) = "3FA62A245A10076D1" '56KEY 9500L0P-432C30E  Serial: 0286103 Arrived 3/18/08
    aryRomSerial(0, 85) = "6BC353746910076D1" '43KEY 9500L0P-412C30E  Serial: 0306767 Arrived 3/18/08
    aryRomSerial(0, 86) = "6B6353846C10076D1" '35Key 9500L0P-422C30E  Serial: 0307458  Roan Mtn, TN
    aryRomSerial(0, 87) = "6BF353946010076D1" '35Key 9500L0P-422C30E  Serial:  0307464 'Currently MNoland's Unit
    aryRomSerial(0, 88) = "6b8353946710076d1" '9500L0P-422c30e (35 key)  S/n 0307465"
    aryRomSerial(0, 89) = "6bd353946010076d1" '9500L0P-422c30e (35 key) s/n 0307457"
    aryRomSerial(0, 90) = "6b0353746f10076d1" '9500L0P-422c30e (35 key) s/n 0306760"
    aryRomSerial(0, 91) = "6bc353746d10076d1" '9500L0P-422c30e (35 key) s/n 0307456"
    aryRomSerial(0, 92) = "6b4353746910076d1" '9500L0P-422c30e (35 key) s/n 0306574"
    aryRomSerial(0, 93) = "6b2353746010076d1" '9500L0P-422c30e (35 key) s/n 0307462"
    aryRomSerial(0, 94) = "6be353946a10076d1" '9500L0P-422c30e (35 key) s/n 0307461"
    aryRomSerial(0, 95) = "6bc353746e10076d1" '9500L0P-422c30e (35 key) s/n 0306582"
    aryRomSerial(0, 96) = "6b2353746a10076d1" '9500L0P-422c30e (35 key) s/n 0306572"
    aryRomSerial(0, 97) = "6bd353946810076d1" '9500L0P-422c30e (35 key) s/n 0307459"
    aryRomSerial(0, 98) = "6bf353746810076d1" '9500L0P-422c30e (35 key) s/n 0306575"
    aryRomSerial(0, 99) = "3f662a245010076d1" '9500L0P-432C30E (56 key) s/n 0286115"
    aryRomSerial(0, 100) = "DUPLICATED" '9500L0P-432C30E (56 key) s/n 0286105"  HH 1084 'Coastal Spartansburg Receiving.  Model#: 9500L0P-432C30E  Serial#: 286105  Manu: 10/26/2007  Install 9/23/2008  HH1084
    aryRomSerial(0, 101) = "3f262a245c10076d1" '9500L0P-432C30E (56 key) s/n 0286122" 'Coastal Spartansburg Receiving.  Model#: 9500L0P-432C30E  Serial#: 286122  Manu: 10/26/2007  Install 9/23/2008  HH1078
    aryRomSerial(0, 102) = "3f162a245210076d1" '9500L0P-432C30E (56 key) s/n 0286117"
    aryRomSerial(0, 103) = "3f062a245a10076d1" '9500L0P-432C30E (56 key) s/n 0286120"
    aryRomSerial(0, 104) = "6bb353746110076d1" '9500L0P-412C30E (43 key) s/n 0306547"
    aryRomSerial(0, 105) = "6be353746310076d1" '9500L0P-412C30E (43 key) s/n 0306563"
    aryRomSerial(0, 106) = "362266e46110076d1" '9500l0p-432c30e  s/n 0317548
    aryRomSerial(0, 107) = "36a266d46010076d1" '9500l0p-432c30e  s/n 0317546
    aryRomSerial(0, 108) = "369266d46f10076d1" '9500l0p-432c30e  s/n 0317547
    aryRomSerial(0, 109) = "361266d46310076d1" '9500l0p-432c30e  s/n 0317549
    aryRomSerial(0, 110) = "367266d46310076d1" '9500l0p-432c30e  s/n 0317552
    aryRomSerial(0, 111) = "36f266d46910076d1" '9500l0p-432c30e  s/n 0317550
    aryRomSerial(0, 112) = "36e266c46e10076d1" '9500l0p-432c30e  s/n 0317513
    aryRomSerial(0, 113) = "36b266c46e10076d1" '9500l0p-432c30e  s/n 0317665
    aryRomSerial(0, 114) = "362266d46110076d1" '9500l0p-432c30e  s/n 0317664
    aryRomSerial(0, 115) = "36f266e46c10076d1" '9500l0p-432c30e  s/n 0317517
    aryRomSerial(0, 116) = "362266e46710076d1" '9500l0p-432c30e  s/n 0317545
    aryRomSerial(0, 117) = "0521B5019E4789F11800-0050BF7A60E2" 'Atlanta Hardwoods New Symbol Unit Purchased by Bruce Manalan 08/01/2008
    aryRomSerial(0, 118) = "006E604F30137" '9900
    aryRomSerial(0, 119) = "343326F46410076D1" 'Coastal Spartansburg Chain Tally.  Model#: 9500L0P-422C30E  Serial#: 317566  Manu: 4/9/2008  Install 9/23/2008   HHXXXX
    aryRomSerial(0, 120) = "34F326F46010076D1" 'Coastal Spartansburg Chain Tally.  Model#: 9500L0P-422C30E  Serial#: 317671  Manu: 4/10/2008  Install 9/23/2008  HHXXXX
    aryRomSerial(0, 121) = "91B151646610076D1" 'Walnut Creek Lumber End Tally.  Model#: 9500L0P-121C030E  Serial#: 0275106  Manu: 9/12/2007
    aryRomSerial(0, 122) = "C59736748810076D1" 'Long Range Scanner 9500 Model   Model#: 9501L0P-132C30E  Serial#: 0343049  Manu: 10/13/2008
    aryRomSerial(0, 123) = "347326F46A10076D1" ' 9500 Model   Model#: 9501L0P-422C30E  Serial#: 0317667  Manu: 4/10/2008
    aryRomSerial(0, 124) = "36F266C46B10076D1" ' 9500 56 Key Serial 0317518 model 9500L0P-432C30E Purchased 11/5/08 Manu 4/9/08 Probably will be at Panpac
    aryRomSerial(0, 125) = "366266C46910076D1" ' 9500 56 Key Serial 0317521 model 9500L0P-432C30E Purchased 11/5/08 Manu 4/9/08 Probably will be at Panpac
    aryRomSerial(0, 126) = "850060A0C031977E0" ' Charles Foreman's Ipaq
    aryRomSerial(0, 127) = "883651A46810076D1" ' 9500 Model  Serial: 0286095
    aryRomSerial(0, 128) = "416313544A10076D1" 'Serial # 0203255  Board Serial # 416313544A10076D1  Highway 47
    aryRomSerial(0, 129) = "41A313544910076D1" 'Serial # 0203249  Board Serial # 41A313544910076D1  Highway 47
    
    aryRomSerial(0, 130) = "C57289A48410076D1" 'Serial # 0223857  Board Serial # C57289A48410076D1  Haessly
    
    aryRomSerial(0, 131) = "006E604F30133" 'HH1253 9900 35 Key Serial # 09228D0093 Pallet One Demo as of 9/9/2009
    aryRomSerial(1, 131) = "SCANWEDGE"
    aryRomSerial(2, 131) = "35KEY"
    aryRomSerial(3, 131) = "9900"
    
    aryRomSerial(0, 132) = "006E604F30132" '56 Key 9900 Serial: 09168D0092  HH1244 Going to Spigelmyer on demo with lumber and logs
    aryRomSerial(1, 132) = "SCANWEDGE"
    aryRomSerial(2, 132) = "35KEY"
    aryRomSerial(3, 132) = "9900"

    aryRomSerial(0, 133) = "006E604F30134" '56 Key 9900 Serial: 08331D0044  HH1244 Going to Spigelmyer on demo with lumber and logs
    aryRomSerial(1, 133) = "SCANWEDGE"
    aryRomSerial(2, 133) = "35KEY"
    aryRomSerial(3, 133) = "9900"
    
    aryRomSerial(0, 134) = "006E604F30144"    ' 9900 35 Key - Duplicated with 43 Unit in Lumber System - Spigelmyer Demo
    aryRomSerial(1, 134) = "SCANWEDGE"
    aryRomSerial(2, 134) = "35KEY"
    aryRomSerial(3, 134) = "9900"
    
    aryRomSerial(0, 135) = "006E604F30139"    ' 9900 43 Key - Purchase approximately 8/1/2009
    aryRomSerial(1, 135) = "SCANWEDGE"
    aryRomSerial(2, 135) = "35KEY"
    aryRomSerial(3, 135) = "9900"
    
    aryRomSerial(0, 136) = "D05366244510076D1"    ' 9500 43 Key -
    aryRomSerial(1, 136) = ""
    aryRomSerial(2, 136) = "43KEY"
    aryRomSerial(3, 136) = "9500"
    
    
    aryRomSerial(0, 137) = "006E604F30132" 'Loaner Demo from HHP for Tradeshows
    aryRomSerial(1, 137) = "SCANWEDGE"
    aryRomSerial(2, 137) = "35KEY"
    aryRomSerial(3, 137) = "9900"
    
    aryRomSerial(0, 138) = "006E604F30138" '56 Key 9900 for Sugar Grove has Log and Lumber System On it.
    aryRomSerial(1, 138) = "SCANWEDGE"     'Duplicate with a 35key, (0, 143) going to Gutchess
    aryRomSerial(2, 138) = "35KEY"         ' Set to 35key for the Gutchess HH
    aryRomSerial(3, 138) = "9900"
    
    aryRomSerial(0, 139) = "006E604F30130" '35 Key 9900 for HH2037 Serial: 9900l0p-331200 McCoy #2
    aryRomSerial(1, 139) = "SCANWEDGE"
    aryRomSerial(2, 139) = "35KEY"
    aryRomSerial(3, 139) = "9900"
    
    aryRomSerial(0, 140) = "006E604F30141" '35 Key 9900 for HH2045 Serial: 9900l0p-331200 DEMO
    aryRomSerial(1, 140) = "SCANWEDGE"
    aryRomSerial(2, 140) = "35KEY"
    aryRomSerial(3, 140) = "9900"
    
    aryRomSerial(0, 141) = "006E604F30135" '35 Key 9900 for HH2061 Serial: 10027D0045  9900l0p-331200 DEMO
    aryRomSerial(1, 141) = "SCANWEDGE"
    aryRomSerial(2, 141) = "35KEY"
    aryRomSerial(3, 141) = "9900"
    
    aryRomSerial(0, 141) = "006E604F30135" '35 Key 9900 for HH2066 Serial:  Duplicaed Key Again on 35 key as usual
    aryRomSerial(1, 141) = "SCANWEDGE" 'Not an accident, please leave here.
    aryRomSerial(2, 141) = "35KEY"
    aryRomSerial(3, 141) = "9900"
    
    aryRomSerial(0, 142) = "006E604F30138" '56 Key 9900 for HH2064 Serial:
    aryRomSerial(1, 142) = "SCANWEDGE"
    aryRomSerial(2, 142) = "56KEY"
    aryRomSerial(3, 142) = "9900"
    
    aryRomSerial(0, 142) = "006E604F30138" '35 Key 9900 for HH2065 Serial:10027D0038
    aryRomSerial(1, 142) = "SCANWEDGE"   'duplicate with a 56key from sugar grove (0, 138)
    aryRomSerial(2, 142) = "35KEY"
    aryRomSerial(3, 142) = "9900"
    
'    aryRomSerial(0, 143) = "50F0063006B000000" '35 Key 9900 for HH2065 Serial:10027D0038
'    aryRomSerial(1, 143) = "SCANWEDGE"   'duplicate with a 56key from sugar grove (0, 138)
'    aryRomSerial(2, 143) = "35KEY"
'    aryRomSerial(3, 143) = "9900"

    '*************GUTCHESS - LOG HANDHELDS
    'All 99EX have the same ROM serial, GUTCHESS has 8 handhelds.
    aryRomSerial(0, 144) = "50F0063006B000000HH2127" '34 key 99EX HH2122
    aryRomSerial(1, 144) = "SCANWEDGE"
    aryRomSerial(2, 144) = "34KEY"
    aryRomSerial(3, 144) = "99EX"

    aryRomSerial(0, 145) = "50F0063006B000000HH2126" '34 key 99EX
    aryRomSerial(1, 145) = "SCANWEDGE"
    aryRomSerial(2, 145) = "34KEY"
    aryRomSerial(3, 145) = "99EX"

    aryRomSerial(0, 146) = "50F0063006B000000HH2129" '34 key 99EX
    aryRomSerial(1, 146) = "SCANWEDGE"
    aryRomSerial(2, 146) = "34KEY"
    aryRomSerial(3, 146) = "99EX"

    aryRomSerial(0, 147) = "50F0063006B000000HH2125" '34 key 99EX
    aryRomSerial(1, 147) = "SCANWEDGE"
    aryRomSerial(2, 147) = "34KEY"
    aryRomSerial(3, 147) = "99EX"

    aryRomSerial(0, 148) = "50F0063006B000000HH2128" '34 key 99EX
    aryRomSerial(1, 148) = "SCANWEDGE"
    aryRomSerial(2, 148) = "34KEY"
    aryRomSerial(3, 148) = "99EX"

    aryRomSerial(0, 149) = "50F0063006B000000HH2130" '34 key 99EX
    aryRomSerial(1, 149) = "SCANWEDGE"
    aryRomSerial(2, 149) = "34KEY"
    aryRomSerial(3, 149) = "99EX"

    aryRomSerial(0, 150) = "50F0063006B000000HH2131" '34 key 99EX
    aryRomSerial(1, 150) = "SCANWEDGE"
    aryRomSerial(2, 150) = "34KEY"
    aryRomSerial(3, 150) = "99EX"

    aryRomSerial(0, 151) = "950065049C10076D1" '56 Key Demo Unit 9500 Series HH1105.
    aryRomSerial(1, 151) = "SCANWEDGE"
    aryRomSerial(2, 151) = "56KEY"
    aryRomSerial(3, 151) = "9500"

    aryRomSerial(0, 152) = "880651A46010076D1" '56 Key Demo Unit 9500 Series HH1105.
    aryRomSerial(1, 152) = "SCANWEDGE"
    aryRomSerial(2, 152) = "43KEY"
    aryRomSerial(3, 152) = "9500"
    
    '**********************************************
    
    aryRomSerial(0, 153) = "50F0063006B000000HH2151" '34 key 99EX
    aryRomSerial(1, 153) = "SCANWEDGE"
    aryRomSerial(2, 153) = "34KEY"
    aryRomSerial(3, 153) = "99EX"
    
    
        'Larry Lines At Gutchess
    aryRomSerial(0, 154) = "50F0063006B000000HH2147" '34 key 99EX 2147
    aryRomSerial(1, 154) = "SCANWEDGE"
    aryRomSerial(2, 154) = "34KEY"
    aryRomSerial(3, 154) = "99EX"
    

    aryRomSerial(0, 155) = "50F0063006B000000HH2191" '34 key 99EX 2191 Gutchess
    aryRomSerial(1, 155) = "SCANWEDGE"
    aryRomSerial(2, 155) = "34KEY"
    aryRomSerial(3, 155) = "99EX"
    
    aryRomSerial(0, 156) = "50F0063006B000000HH2253" '34 key 99EX 2253 Gutchess
    aryRomSerial(1, 156) = "SCANWEDGE"
    aryRomSerial(2, 156) = "34KEY"
    aryRomSerial(3, 156) = "99EX"
    
    aryRomSerial(0, 157) = "50F0063006B000000HH2252" '34 key 99EX 2252 Gutchess
    aryRomSerial(1, 157) = "SCANWEDGE"
    aryRomSerial(2, 157) = "34KEY"
    aryRomSerial(3, 157) = "99EX"
    
    '******************Not Licensed Added for new ones
    
    aryRomSerial(0, 158) = "50F0063006B000000HH4058" '34 key 99EX 'HH2275 Pallet One 8/6/12
    aryRomSerial(1, 158) = "SCANWEDGE"
    aryRomSerial(2, 158) = "34KEY"
    aryRomSerial(3, 158) = "99EX"
    
    aryRomSerial(0, 159) = "50F0063006B000000HH4059" '34 key 99EX 'HH2274 Pallet One 8/6/12
    aryRomSerial(1, 159) = "SCANWEDGE"
    aryRomSerial(2, 159) = "34KEY"
    aryRomSerial(3, 159) = "99EX"
    
    aryRomSerial(0, 160) = "50F0063006B000000HH4060" '34 key 99EX Extra
    aryRomSerial(1, 160) = "SCANWEDGE"
    aryRomSerial(2, 160) = "34KEY"
    aryRomSerial(3, 160) = "99EX"
    
    aryRomSerial(0, 161) = "50F0063006B000000HH4061" '34 key 99EX Extra
    aryRomSerial(1, 161) = "SCANWEDGE"
    aryRomSerial(2, 161) = "34KEY"
    aryRomSerial(3, 161) = "99EX"
    
    aryRomSerial(0, 162) = "50F0063006B000000HH4062" '34 key 99EX Extra
    aryRomSerial(1, 162) = "SCANWEDGE"
    aryRomSerial(2, 162) = "34KEY"
    aryRomSerial(3, 162) = "99EX"
    
    aryRomSerial(0, 163) = "50F0063006B000000HH4063" '34 key 99EX Extra
    aryRomSerial(1, 163) = "SCANWEDGE"
    aryRomSerial(2, 163) = "34KEY"
    aryRomSerial(3, 163) = "99EX"
    
    aryRomSerial(0, 164) = "50F0063006B000000HH4064" '34 key 99EX Extra
    aryRomSerial(1, 164) = "SCANWEDGE"
    aryRomSerial(2, 164) = "34KEY"
    aryRomSerial(3, 164) = "99EX"
    
    aryRomSerial(0, 165) = "50F0063006B000000HH4065" '34 key 99EX Extra
    aryRomSerial(1, 165) = "SCANWEDGE"
    aryRomSerial(2, 165) = "34KEY"
    aryRomSerial(3, 165) = "99EX"
    
    aryRomSerial(0, 166) = "50F0063006B000000HH4066" '34 key 99EX Extra
    aryRomSerial(1, 166) = "SCANWEDGE"
    aryRomSerial(2, 166) = "34KEY"
    aryRomSerial(3, 166) = "99EX"
    
    aryRomSerial(0, 167) = "50F0063006B000000HH4067" '34 key 99EX Extra
    aryRomSerial(1, 167) = "SCANWEDGE"
    aryRomSerial(2, 167) = "34KEY"
    aryRomSerial(3, 167) = "99EX"
    
    aryRomSerial(0, 168) = "50F0063006B000000HH4068" '34 key 99EX Extra
    aryRomSerial(1, 168) = "SCANWEDGE"
    aryRomSerial(2, 168) = "34KEY"
    aryRomSerial(3, 168) = "99EX"
    
    aryRomSerial(0, 169) = "50F0063006B000000HH4069" '34 key 99EX Extra
    aryRomSerial(1, 169) = "SCANWEDGE"
    aryRomSerial(2, 169) = "34KEY"
    aryRomSerial(3, 169) = "99EX"
    
    aryRomSerial(0, 170) = "50F0063006B000000HH2234" '34 key Chuck CW Logs
    aryRomSerial(1, 170) = "SCANWEDGE"
    aryRomSerial(2, 170) = "34KEY"
    aryRomSerial(3, 170) = "99EX"
    
    aryRomSerial(0, 171) = "796925344A10076D1" 'HIGHWAY 47
    aryRomSerial(1, 171) = "9500"
    '
    
    aryRomSerial(0, 172) = "50F0063006B000000HH2223" 'HH2223 serial 13104D8297 New Taylor 34 Key 99EX 7/1/13 setting up for buyer with HP 100, new version of log tally for v2 report printing
    aryRomSerial(1, 172) = "SCANWEDGE"
    aryRomSerial(2, 172) = "34KEY"
    aryRomSerial(3, 172) = "99EX"
    
    aryRomSerial(0, 173) = "50F0063006B00000011179D0129" 'HH2177 serial 11179D0129 New Quality (MI) Handheld for Logs
    aryRomSerial(1, 173) = "SCANWEDGE"
    aryRomSerial(2, 173) = "55KEY"
    aryRomSerial(3, 173) = "99EX"
    
    aryRomSerial(0, 174) = "50F0063006B00000013070D807D" 'HH2325 serial 13070D807D New Quality (MI) Handheld for Logs
    aryRomSerial(1, 174) = "SCANWEDGE"
    aryRomSerial(2, 174) = "34KEY"
    aryRomSerial(3, 174) = "99EX"
    
    aryRomSerial(0, 175) = "50F0063006B00000012188D01F8" 'HH2306 serial 12188D01F8 Plum Creek Handheld for Logs
    aryRomSerial(1, 175) = "SCANWEDGE"
    aryRomSerial(2, 175) = "34KEY"
    aryRomSerial(3, 175) = "99EX"
    
    aryRomSerial(0, 176) = "50F0063006B000000HH7321" 'HH7321 serial 13150D0C4 Tri State Handheld for Logs
    aryRomSerial(1, 176) = "SCANWEDGE"
    aryRomSerial(2, 176) = "34KEY"
    aryRomSerial(3, 176) = "99EX"
    
    aryRomSerial(0, 177) = "50F0063006B00000013113D821D" 'HH2317 Serial 13113D821D 55 Key- Fredonia Returned From RMA
    aryRomSerial(1, 177) = "SCANWEDGE"
    aryRomSerial(2, 177) = "55KEY"
    aryRomSerial(3, 177) = "99EX"
    
    '
    
    aryRomSerial(0, 178) = "50F0063006B00000013150D80CE" 'HH 7338 Serial 13150D80CE 34 Key- Quality MI
    aryRomSerial(1, 178) = "SCANWEDGE"
    aryRomSerial(2, 178) = "34KEY"
    aryRomSerial(3, 178) = "99EX"
    
    aryRomSerial(0, 179) = "50F0063006B00000013150D8162" 'HH 7349 Serial 13150D8162 34 Key- C.C. COOK & SON
    aryRomSerial(1, 179) = "SCANWEDGE"
    aryRomSerial(2, 179) = "34KEY"
    aryRomSerial(3, 179) = "99EX"
    
    aryRomSerial(0, 180) = "50F0063006B00000011361D0258" 'HH 2244 Serial 11361D0258 'Requested by Brewster 2014-01-28
    aryRomSerial(1, 180) = "SCANWEDGE"
    aryRomSerial(2, 180) = "34KEY"
    aryRomSerial(3, 180) = "99EX"
    
    aryRomSerial(0, 181) = "50F0063006B00000013086D8083" 'HH 7360 Serial 13086D8083 'PreLicense
    aryRomSerial(1, 181) = "SCANWEDGE"
    aryRomSerial(2, 181) = "34KEY"
    aryRomSerial(3, 181) = "99EX"
    
    aryRomSerial(0, 182) = "50F0063006B00000013085D839A" 'HH 7361 Serial 13085D839A  'PreLicense
    aryRomSerial(1, 182) = "SCANWEDGE"
    aryRomSerial(2, 182) = "34KEY"
    aryRomSerial(3, 182) = "99EX"
    
    aryRomSerial(0, 183) = "50F0063006B00000014019D830D"  'HH Serial 14019D830D  HH7364 55 KEY SUGAR GROVE 2/26/2014   99EX-Licensed for Log and Lumber
    aryRomSerial(1, 183) = "SCANWEDGE"
    aryRomSerial(2, 183) = "55KEY"
    aryRomSerial(3, 183) = "99EX"
    
    aryRomSerial(0, 184) = "50F0063006B00000013086D8051"  'HH Serial 13086D8051  HH7366 34 Key Dimension Hardwood Veneers
    aryRomSerial(1, 184) = "SCANWEDGE"
    aryRomSerial(2, 184) = "34KEY"
    aryRomSerial(3, 184) = "99EX"
    
    aryRomSerial(0, 185) = "50F0063006B00000013086D805E"  'HH Serial 13086D805E  HH7359 34 Key Baxter Lumber
    aryRomSerial(1, 185) = "SCANWEDGE"
    aryRomSerial(2, 185) = "34KEY"
    aryRomSerial(3, 185) = "99EX"
    
    aryRomSerial(0, 186) = "50F0063006B00000013086D80A4"  'HH Serial 13086D80A4  HH7358 34 Key Besse Lumber
    aryRomSerial(1, 186) = "SCANWEDGE"
    aryRomSerial(2, 186) = "34KEY"
    aryRomSerial(3, 186) = "99EX"
    

    aryRomSerial(0, 187) = "50F0063006B00000014104D81A0"  'HH Serial 14104D81A0  HH7383 JM Wood Products (StoneysHardwoods' AKA)
    aryRomSerial(1, 187) = "SCANWEDGE"
    aryRomSerial(2, 187) = "34KEY"
    aryRomSerial(3, 187) = "99EX"
    
    
    aryRomSerial(0, 188) = "50F0063006B00000014104D8194"  'HH Serial 14104D8194  HH7375 34 Key Horizon Wood Products
    aryRomSerial(1, 188) = "SCANWEDGE"
    aryRomSerial(2, 188) = "34KEY"
    aryRomSerial(3, 188) = "99EX"

    aryRomSerial(0, 189) = "50F0063006B00000014104D820E"  'HH Serial 14104D820E  HH7378 34 Key Horizon Wood Products
    aryRomSerial(1, 189) = "SCANWEDGE"
    aryRomSerial(2, 189) = "34KEY"
    aryRomSerial(3, 189) = "99EX"


    aryRomSerial(0, 190) = "50F0063006B00000014105D802D"  'HH Serial 14105D802D  HH7379 34 Key Horizon Wood Products
    aryRomSerial(1, 190) = "SCANWEDGE"
    aryRomSerial(2, 190) = "34KEY"
    aryRomSerial(3, 190) = "99EX"

    aryRomSerial(0, 191) = "50F0063006B00000014105D802F"  'HH Serial 14105D802F  HH7380 34 Key Horizon Wood Products
    aryRomSerial(1, 191) = "SCANWEDGE"
    aryRomSerial(2, 191) = "34KEY"
    aryRomSerial(3, 191) = "99EX"


    aryRomSerial(0, 300) = "50F0063006B00000014105D8029"  'HH Serial 14105D8029  HH7376 34 Key Trico Lumber
    aryRomSerial(1, 300) = "SCANWEDGE"
    aryRomSerial(2, 300) = "34KEY"
    aryRomSerial(3, 300) = "99EX"

    aryRomSerial(0, 301) = "50F0063006B00000014104D8176"  'HH Serial 14104D8176  HH7381 34 Key Wilson Forest  Products (Steve Bee-Log Buyer)
    aryRomSerial(1, 301) = "SCANWEDGE"
    aryRomSerial(2, 301) = "34KEY"
    aryRomSerial(3, 301) = "99EX"

    aryRomSerial(0, 302) = "50F0063006B00000014112D823B"  'HH Serial 14112D823B  HH7409 34 Key Kuhl LUmber 9/2/14
    aryRomSerial(1, 302) = "SCANWEDGE"
    aryRomSerial(2, 302) = "34KEY"
    aryRomSerial(3, 302) = "99EX"

    aryRomSerial(0, 303) = "50F0063006B00000014211D83FA"  'HH Serial 14211D83FA  99EXL01-0C212SE   HH7408 34 Key ATCO Log 9/9/14
    aryRomSerial(1, 303) = "SCANWEDGE"
    aryRomSerial(2, 303) = "34KEY"
    aryRomSerial(3, 303) = "99EX"
    
    aryRomSerial(0, 304) = "50F0063006B00000014112D8237"  'HH Serial  14112D8237  99EXL01-0C212SE HH7407 34 Key ATCO Log 9/9/14
    aryRomSerial(1, 304) = "SCANWEDGE"
    aryRomSerial(2, 304) = "34KEY"
    aryRomSerial(3, 304) = "99EX"
    
    
    aryRomSerial(0, 305) = "50F0063006B00000014112D8245"  'HH Serial 14112D8246  99EXL01-0C212SE  HH7406 34 Key ATCO Log 9/9/14
    aryRomSerial(1, 305) = "SCANWEDGE"
    aryRomSerial(2, 305) = "34KEY"
    aryRomSerial(3, 305) = "99EX"
    
    aryRomSerial(0, 306) = "50F0063006B00000014171D8225"  'HH Serial 14171D8225  99EXL01-0C212SE  HH7398 34 Key Angel Forest Products 9/10/14
    aryRomSerial(1, 306) = "SCANWEDGE"
    aryRomSerial(2, 306) = "34KEY"
    aryRomSerial(3, 306) = "99EX"
    
     
    aryRomSerial(0, 307) = "50F0063006B00000012005D00A3" 'HH Serial  12005D00A3   HH7414   RAM Timber Cruise/Log Handhelds with GPS
    aryRomSerial(1, 307) = "SCANWEDGE"
    aryRomSerial(2, 307) = "34KEY"
    aryRomSerial(3, 307) = "99EX"
    
    
    aryRomSerial(0, 308) = "50F0063006B00000012005D01BF" 'HH Serial  12005D01BF   HH7416   RAM Timber Cruise/Log Handhelds with GPS
    aryRomSerial(1, 308) = "SCANWEDGE"
    aryRomSerial(2, 308) = "34KEY"
    aryRomSerial(3, 308) = "99EX"
    
    
    aryRomSerial(0, 309) = "50F0063006B00000012115D0124" 'HH Serial  12115D0124   HH7417   RAM Timber Cruise/Log Handhelds with GPS
    aryRomSerial(1, 309) = "SCANWEDGE"
    aryRomSerial(2, 309) = "34KEY"
    aryRomSerial(3, 309) = "99EX"
    
    
    aryRomSerial(0, 310) = "50F0063006B00000012103D024F" 'HH Serial  12103D024F   HH7415   RAM Timber Cruise/Log Handhelds with GPS
    aryRomSerial(1, 310) = "SCANWEDGE"
    aryRomSerial(2, 310) = "34KEY"
    aryRomSerial(3, 310) = "99EX"
    
    
    aryRomSerial(0, 311) = "50F0063006B00000014186D82C0"  'HH Serial 14186D82C0  HH7418 34 Key J.M. Longyear
    aryRomSerial(1, 311) = "SCANWEDGE"
    aryRomSerial(2, 311) = "34KEY"
    aryRomSerial(3, 311) = "99EX"
 
    aryRomSerial(0, 312) = "50F0063006B00000014186D8366"  'HH Serial 14186D8366  HH7419 34 Key J.M. Longyear
    aryRomSerial(1, 312) = "SCANWEDGE"
    aryRomSerial(2, 312) = "34KEY"
    aryRomSerial(3, 312) = "99EX"
 
 
    aryRomSerial(0, 313) = "50F0063006B00000014186D8324"  'HH Serial 14186D8324  HH7420 34 Key Horizon Wood Products
    aryRomSerial(1, 313) = "SCANWEDGE"
    aryRomSerial(2, 313) = "34KEY"
    aryRomSerial(3, 313) = "99EX"
    
    aryRomSerial(0, 314) = "50F0063006B00000014186D8297"  'HH Serial 14186D8297  HH7411 34 Key Glatfelter - Piketon
    aryRomSerial(1, 314) = "SCANWEDGE"
    aryRomSerial(2, 314) = "34KEY"
    aryRomSerial(3, 314) = "99EX"
 
    aryRomSerial(0, 315) = "50F0063006B00000014115D80D6"  'HH Serial 14115D80D6  HH7412 34 Key Glatfelter - Piketon
    aryRomSerial(1, 315) = "SCANWEDGE"
    aryRomSerial(2, 315) = "34KEY"
    aryRomSerial(3, 315) = "99EX"
    
    aryRomSerial(0, 316) = "50F0063006B00000013150D80C5"  'HH Serial 13150D80C5  HH7331 34 Key CLC Hardwoods - Handheld Licensed On RMA/Repair Service, apparently not when originally sold
    aryRomSerial(1, 316) = "SCANWEDGE"
    aryRomSerial(2, 316) = "34KEY"
    aryRomSerial(3, 316) = "99EX"
    
    aryRomSerial(0, 317) = "50F0063006B00000011329D028E"  'HH Serial 11329D028E HH2235 55 Key Quality Hardwoods, Inc (MI) - Sold a while ago on 8/22/2012
    aryRomSerial(1, 317) = "SCANWEDGE"
    aryRomSerial(2, 317) = "55KEY"
    aryRomSerial(3, 317) = "99EX"
    
    aryRomSerial(0, 318) = "50F0063006B00000014258D8041"  'HH Serial 14258D8041 HH7428 34 Key Log & Lumber Licensed - Gilkey Lumber 11/17/14
    aryRomSerial(1, 318) = "SCANWEDGE"
    aryRomSerial(2, 318) = "34KEY"
    aryRomSerial(3, 318) = "99EX"
    
    aryRomSerial(0, 319) = "50F0063006B00000014258D80EA"  'HH Serial 14258D80EA  HH7424 34 Key McKay Hardwoods
    aryRomSerial(1, 319) = "SCANWEDGE"
    aryRomSerial(2, 319) = "34KEY"
    aryRomSerial(3, 319) = "99EX"
    
    
    aryRomSerial(0, 320) = "50F0063006B00000014258D8135"  'HH Serial 14258D8135 HH7443 34 Key Yoder Lumber Co. mLOGS 99ex 34kHH7443 14258D8135
    aryRomSerial(1, 320) = "SCANWEDGE"
    aryRomSerial(2, 320) = "34KEY"
    aryRomSerial(3, 320) = "99EX"
    
    aryRomSerial(0, 321) = "50F0063006B00000013151D81DD"  'HH Serial 13151D81DD HH7339 34 Key Bryant and Young
    aryRomSerial(1, 321) = "SCANWEDGE"
    aryRomSerial(2, 321) = "34KEY"
    aryRomSerial(3, 321) = "99EX"
    
    aryRomSerial(0, 322) = "50F0063006B00000014258D8041"  'HH Serial 14258D8041 HH7428 34 Key Gilkey Lumber
    aryRomSerial(1, 322) = "SCANWEDGE"
    aryRomSerial(2, 322) = "34KEY"
    aryRomSerial(3, 322) = "99EX"
   
    aryRomSerial(0, 323) = "50F0063006B00000014258D80B3"  'HH Serial 14258D80B3 HH7425 34 Key Quality of MI
    aryRomSerial(1, 323) = "SCANWEDGE"
    aryRomSerial(2, 323) = "34KEY"
    aryRomSerial(3, 323) = "99EX"
   
    aryRomSerial(0, 324) = "50F0063006B00000014253D8359"  'HH Serial 14253D8359 HH7465 34 Key Meherrin River
    aryRomSerial(1, 324) = "SCANWEDGE"
    aryRomSerial(2, 324) = "34KEY"
    aryRomSerial(3, 324) = "99EX"
   
    aryRomSerial(0, 325) = "50F0063006B00000014258D8038"  'HH Serial 14258D8038 HH7426 34 Key Mark Depp - Independent Grader 4/7/15
    aryRomSerial(1, 325) = "SCANWEDGE"
    aryRomSerial(2, 325) = "34KEY"
    aryRomSerial(3, 325) = "99EX"
    aryRomSerial(4, 325) = "LOGTALLYONLY-RESTRICTED"
    
    aryRomSerial(0, 326) = "50F0063006B00000014359D8570"  'HH Serial 14359D8570 HH7543 34 Key J.M. Longyear 4/21/15
    aryRomSerial(1, 326) = "SCANWEDGE"
    aryRomSerial(2, 326) = "34KEY"
    aryRomSerial(3, 326) = "99EX"
    
    aryRomSerial(0, 327) = "50F0063006B00000014359D8375"  'HH Serial 14359D8375 HH7545 34 Key J.M. Longyear 4/21/15
    aryRomSerial(1, 327) = "SCANWEDGE"
    aryRomSerial(2, 327) = "34KEY"
    aryRomSerial(3, 327) = "99EX"
    
    aryRomSerial(0, 328) = "50F0063006B00000014359D8292"  'HH Serial 14359D8292 HH7542 34 Key Quality Hardwood Products 4/21/15
    aryRomSerial(1, 328) = "SCANWEDGE"
    aryRomSerial(2, 328) = "34KEY"
    aryRomSerial(3, 328) = "99EX"
    
    aryRomSerial(0, 329) = "50F0063006B00000014256D807D"  'HH Serial 14256D807D HH7559 34 Key Yoder Lumber Company 5/12/15
    aryRomSerial(1, 329) = "SCANWEDGE"
    aryRomSerial(2, 329) = "34KEY"
    aryRomSerial(3, 329) = "99EX"
    
    aryRomSerial(0, 330) = "50F0063006B00000014254D849C"  'HH Serial 14254D849C HH7561 34 Key Yoder Lumber Company 5/12/15
    aryRomSerial(1, 330) = "SCANWEDGE"
    aryRomSerial(2, 330) = "34KEY"
    aryRomSerial(3, 330) = "99EX"
    
    aryRomSerial(0, 331) = "50F0063006B00000014255D8389"  'HH Serial 14255D8389 HHD7573 34 Key Andis Logging 5/11/15
    aryRomSerial(1, 331) = "SCANWEDGE"
    aryRomSerial(2, 331) = "34KEY"
    aryRomSerial(3, 331) = "99EX"
    
    aryRomSerial(0, 332) = "50F0063006B00000015008D8519"  'HH Serial 15008D8519 HHL7474 34 Key Two Rivers  5/22/15
    aryRomSerial(1, 332) = "SCANWEDGE"
    aryRomSerial(2, 332) = "34KEY"
    aryRomSerial(3, 332) = "99EX"
    
    'aryRomSerial(0, 333) = "50F0063006B00000014106D8214"  'HH7390 Serial 14106D8214 eLIMBS Loaner Handheld - 7390 - Was at bolke veneer for a long time, now being added to loaner inventory and sent on a demo 6/16/15
    'aryRomSerial(1, 333) = "SCANWEDGE"
    'aryRomSerial(2, 333) = "34KEY"
    'aryRomSerial(3, 333) = "99EX"
    
    'aryRomSerial(0, 334) = "50F0063006B00000014106D8214"  'HH7586 Serial 14106D8214 eLIMBS Loaner Handheld - HH7586 - Being sent as a Demo
    'aryRomSerial(1, 334) = "SCANWEDGE"
    'aryRomSerial(2, 334) = "34KEY"
    'aryRomSerial(3, 334) = "99EX"
    
    aryRomSerial(0, 335) = "50F0063006B00000014300D8154"  'HH7602 Serial 14300D8154 eLIMBS Brian Wilson 99EX Log Replacement for 9900.
    aryRomSerial(1, 335) = "SCANWEDGE"                    'Licensed  - 7/1/15
    aryRomSerial(2, 335) = "34KEY"
    aryRomSerial(3, 335) = "99EX"
    
    aryRomSerial(0, 336) = "50F0063006B00000014326D8270"  'HHL7446  Serial: 14326D8270 eLIMBS Meherrin River Setup 6/12/2015 - but somehow not licensed during setup
    aryRomSerial(1, 336) = "SCANWEDGE"                    'Licensed  - 7/1/15
    aryRomSerial(2, 336) = "55KEY"
    aryRomSerial(3, 336) = "99EX"
    
    aryRomSerial(0, 337) = "50F0063006B00000014326D8117"  'HHL7447  Serial: 14326D8117 eLIMBS Meherrin River Setup 6/12/2015 - but somehow not licensed during setup
    aryRomSerial(1, 337) = "SCANWEDGE"                     'Licensed  - 7/1/15
    aryRomSerial(2, 337) = "55KEY"
    aryRomSerial(3, 337) = "99EX"
    
    aryRomSerial(0, 338) = "50F0063006B00000014326D814B"  'HHL7448  Serial: 14326D814B eLIMBS Meherrin River Setup 6/12/2015 - but somehow not licensed during setup
    aryRomSerial(1, 338) = "SCANWEDGE"                    'Licensed 7/1/15 - shipped 6/12/15 - no shipping verification signature.
    aryRomSerial(2, 338) = "55KEY"                        'Re-Licensed 2/11/16 - was licensed with 14326D8143, but should have been 14326D814B.  RM
    aryRomSerial(3, 338) = "99EX"
    
    aryRomSerial(0, 339) = "50F0063006B00000014301D80C4"  'HH Serial 14301D80C4 HH7585 34 Key Andis Logging 7/1/15 license by PCC
    aryRomSerial(1, 339) = "SCANWEDGE"
    aryRomSerial(2, 339) = "34KEY"
    aryRomSerial(3, 339) = "99EX"
    
    aryRomSerial(0, 340) = "50F0063006B00000012189D0143"  'HH Serial 12189D0143 HHL2285 34 Key McGee Lumber Co., Inc.
    aryRomSerial(1, 340) = "SCANWEDGE"
    aryRomSerial(2, 340) = "34KEY"
    aryRomSerial(3, 340) = "99EX"
       
    aryRomSerial(0, 341) = "50F0063006B00000014255D83B8"  'HH Serial 14255D83B8 HH7603 34 Key Two Rivers
    aryRomSerial(1, 341) = "SCANWEDGE"
    aryRomSerial(2, 341) = "34KEY"
    aryRomSerial(3, 341) = "99EX"
    
    aryRomSerial(0, 342) = "50F0063006B00000014255D83E5"  'HH Serial 14255D83E5 HH7604 34 Springfield Hardwood
    aryRomSerial(1, 342) = "SCANWEDGE"
    aryRomSerial(2, 342) = "34KEY"
    aryRomSerial(3, 342) = "99EX"
   
    aryRomSerial(0, 343) = "50F0063006B00000014253D8330"  'HH Serial 14253D8330 HH7421 34 Quality Hardwood Products
    aryRomSerial(1, 343) = "SCANWEDGE"
    aryRomSerial(2, 343) = "34KEY"
    aryRomSerial(3, 343) = "99EX"
    
    aryRomSerial(0, 344) = "50F0063006B00000014281D852C"  'HH Serial 14281D852C HH7506 34 Quality Hardwood Products
    aryRomSerial(1, 344) = "SCANWEDGE"
    aryRomSerial(2, 344) = "34KEY"
    aryRomSerial(3, 344) = "99EX"
    
    aryRomSerial(0, 345) = "50F0063006B00000014281D82A8"  'HH Serial 14281D82A8 HH7507 34 Quality Hardwood Products
    aryRomSerial(1, 345) = "SCANWEDGE"
    aryRomSerial(2, 345) = "34KEY"
    aryRomSerial(3, 345) = "99EX"
    
    aryRomSerial(0, 346) = "50F0063006B00000014359D8435"  'HH Serial 14359D8435 HH7544 34 Quality Hardwood Products
    aryRomSerial(1, 346) = "SCANWEDGE"
    aryRomSerial(2, 346) = "34KEY"
    aryRomSerial(3, 346) = "99EX"
   
    aryRomSerial(0, 347) = "50F0063006B00000014281D82C0"  'HH Serial 14281D82C0 HH7628 34 Kirkham / Maley Wertz
    aryRomSerial(1, 347) = "SCANWEDGE"
    aryRomSerial(2, 347) = "34KEY"
    aryRomSerial(3, 347) = "99EX"
    
    aryRomSerial(0, 348) = "50F0063006B00000014281D8315"  'HH Serial 14281D8315 HH7629 34 Kirkham / Maley Wertz
    aryRomSerial(1, 348) = "SCANWEDGE"
    aryRomSerial(2, 348) = "34KEY"
    aryRomSerial(3, 348) = "99EX"
   
    aryRomSerial(0, 349) = "50F0063006B00000015008D857F"  'HH Serial 15008D857F  HHL7475 34 8/24/2015 - PCC Licensed Loaner(Honeywell Loaner) per Amy They Keeping Keweenaw Land Association - for Nic Setup
    aryRomSerial(1, 349) = "SCANWEDGE"
    aryRomSerial(2, 349) = "34KEY"
    aryRomSerial(3, 349) = "99EX"
    
    aryRomSerial(0, 350) = "50F0063006B00000015008D837A"  'HH Serial 15008D837A  HHL7471 34 9/4/2015 - NAJ Licensed Loaner(Honeywell Loaner) per Amy They Keeping Megnin Cooperage Mills - for Nic Setup
    aryRomSerial(1, 350) = "SCANWEDGE"
    aryRomSerial(2, 350) = "34KEY"
    aryRomSerial(3, 350) = "99EX"
    
    aryRomSerial(0, 351) = "50F0063006B00000014281D832A"  'HH Serial 14281D832A  HH7634 34 9/30/2015 - NAJ Loggers Inc.
    aryRomSerial(1, 351) = "SCANWEDGE"
    aryRomSerial(2, 351) = "34KEY"
    aryRomSerial(3, 351) = "99EX"
    
    'aryRomSerial(0, 352) = "50F0063006B00000014106D8214"  'HH Serial 14106D8214  HH7390 34 10/2/2015 - NAJ Robinson Lumber
    'aryRomSerial(1, 352) = "SCANWEDGE"
    'aryRomSerial(2, 352) = "34KEY"
    'aryRomSerial(3, 352) = "99EX"
    
    aryRomSerial(0, 353) = "50F0063006B00000014351D82D3"  'HH Serial 14351D82D3  HH7656 34 10/2/2015 - NAJ Quality IN Ben Vineyard to replace stolen one
    aryRomSerial(1, 353) = "SCANWEDGE"
    aryRomSerial(2, 353) = "34KEY"
    aryRomSerial(3, 353) = "99EX"
    
    aryRomSerial(0, 354) = "50F0063006B00000013086D80A4"  'HH Serial 13086D80A4  HH7358 34 10/23/2015 - NAJ Sylvan Timber Clearing
    aryRomSerial(1, 354) = "SCANWEDGE"
    aryRomSerial(2, 354) = "34KEY"
    aryRomSerial(3, 354) = "99EX"
    
    aryRomSerial(0, 355) = "50F0063006B00000015169D867C"  'HH Serial 15169D867C  HH7657 34 10/29/2015 - NAJ Gift Lumber
    aryRomSerial(1, 355) = "SCANWEDGE"
    aryRomSerial(2, 355) = "34KEY"
    aryRomSerial(3, 355) = "99EX"
    
    aryRomSerial(0, 356) = "50F0063006B00000014359D8292"  'HH Serial 14359D8292  HH7542 34 12/03/2015 - NAJ Green Ridge
    aryRomSerial(1, 356) = "SCANWEDGE"
    aryRomSerial(2, 356) = "34KEY"
    aryRomSerial(3, 356) = "99EX"
    
    aryRomSerial(0, 357) = "50F0063006B00000014106D8214"  'HH Serial 14106D8214  HH7390 34 12/4/2015 - NAJ - Prime Supply
    aryRomSerial(1, 357) = "SCANWEDGE"
    aryRomSerial(2, 357) = "34KEY"
    aryRomSerial(3, 357) = "99EX"
    
    aryRomSerial(0, 358) = "50F0063006B00000012129D0037"  'HH Serial 12129D0037  HH2306 34 12/09/2015 - NAJ - Rock Lumber
    aryRomSerial(1, 358) = "SCANWEDGE"
    aryRomSerial(2, 358) = "34KEY"
    aryRomSerial(3, 358) = "99EX"
    
    aryRomSerial(0, 359) = "50F0063006B00000014362D80F6"  'HH Serial 14362D80F6  HH7671 34 1/11/2016 - NAJ - Green Ridge new HH
    aryRomSerial(1, 359) = "SCANWEDGE"
    aryRomSerial(2, 359) = "34KEY"
    aryRomSerial(3, 359) = "99EX"
    
    aryRomSerial(0, 360) = "50F0063006B00000014362D80EC"  'HH Serial 14362D80EC  HH7670 34 1/11/2016 - NAJ - T&G Lumber
    aryRomSerial(1, 360) = "SCANWEDGE"
    aryRomSerial(2, 360) = "34KEY"
    aryRomSerial(3, 360) = "99EX"
    
    aryRomSerial(0, 361) = "0B400D542A10076D1"    ' 9500 43 Key - S&J Lumber - Alan bought these from somewhere else and we do not support the hardware
    aryRomSerial(1, 361) = ""
    aryRomSerial(2, 361) = "43KEY"
    aryRomSerial(3, 361) = "9500"
    
    aryRomSerial(0, 362) = "0286150"    ' 9500 43 HH119 Key - S&J Lumber - Alan bought these from somewhere else and we do not support the hardware
    aryRomSerial(1, 362) = ""
    aryRomSerial(2, 362) = "43KEY"
    aryRomSerial(3, 362) = "9500"
    
    aryRomSerial(0, 363) = "50F0063006B00000014103D8188"    'HH Serial 14103D8188  HH7679 34 1/22/2016 - NAJ - JM Longyear
    aryRomSerial(1, 363) = "SCANWEDGE"
    aryRomSerial(2, 363) = "34KEY"
    aryRomSerial(3, 363) = "99EX"
    
    aryRomSerial(0, 364) = "50F0063006B00000014362D80FA"    'HH Serial 14362D80FA  HH7680 34 1/22/2016 - NAJ - JM Longyear
    aryRomSerial(1, 364) = "SCANWEDGE"
    aryRomSerial(2, 364) = "34KEY"
    aryRomSerial(3, 364) = "99EX"
    
    aryRomSerial(0, 365) = "50F0063006B00000014281D829E"    'HH Serial 14281D829E  HH7684 34 1/27/2016 - NAJ - Ram Forest Products
    aryRomSerial(1, 365) = "SCANWEDGE"
    aryRomSerial(2, 365) = "34KEY"
    aryRomSerial(3, 365) = "99EX"
    
    aryRomSerial(0, 366) = "50F0063006B00000012128D0166"    'HH Serial 12128D0166  HH2307 34 2/01/2016 - NAJ - Kendrick Forest Products
    aryRomSerial(1, 366) = "SCANWEDGE"
    aryRomSerial(2, 366) = "34KEY"
    aryRomSerial(3, 366) = "99EX"
    
    aryRomSerial(0, 367) = "50F0063006B00000014341D8153"    'HH Serial 14341D8153  HH7693 34 10/4/2016 - NAJ - Fredonia Forest Products
    aryRomSerial(1, 367) = "SCANWEDGE"                      'HH Serial 14341D8153  HH7693 34 2/08/2016 - NAJ - Townsend Lumber (Old value)
    aryRomSerial(2, 367) = "34KEY"
    aryRomSerial(3, 367) = "99EX"
    
    aryRomSerial(0, 368) = "50F0063006B00000011361D0259"    'HH Serial 11361D0259  HHD2244 55 2/09/2016 - NAJ - Meherin River - This somehow did not get licensed
    aryRomSerial(1, 368) = "SCANWEDGE"
    aryRomSerial(2, 368) = "55KEY"
    aryRomSerial(3, 368) = "99EX"
    
    aryRomSerial(0, 369) = "50F0063006B00000015147D8389"    'HH Serial 15147D8389  HH7702 34 2/18/2016 - NAJ - Green Ridge
    aryRomSerial(1, 369) = "SCANWEDGE"
    aryRomSerial(2, 369) = "55KEY"
    aryRomSerial(3, 369) = "99EX"
    
    aryRomSerial(0, 370) = "50F0063006B00000011231D0034"    'HH Serial 11231D0034  HH2169 55 2/23/2016 - RM - Krueger (existing mLIMBS HH)
    aryRomSerial(1, 370) = "SCANWEDGE"
    aryRomSerial(2, 370) = "55KEY"
    aryRomSerial(3, 370) = "99EX"
    
    aryRomSerial(0, 371) = "50F0063006B00000014281D82CC"    'HH Serial 14281D82CC  HH7692 34 3/23/2016 - BJP - Rock Hardwoods
    aryRomSerial(1, 371) = "SCANWEDGE"
    aryRomSerial(2, 371) = "34KEY"
    aryRomSerial(3, 371) = "99EX"
    
    aryRomSerial(0, 372) = "6B3353746110076D1"    ' 9500 43 HH1075 0306777 43 Key - Just in case we need a loaner NJ (updated by RM)
    aryRomSerial(1, 372) = ""
    aryRomSerial(2, 372) = "43KEY"
    aryRomSerial(3, 372) = "9500"
    
    aryRomSerial(0, 373) = "50F0063006B00000016089D8278"    'HH Serial 16089D8278  HH7714 34 4/14/2016 - NAJ - Promark Percussion
    aryRomSerial(1, 373) = "SCANWEDGE"
    aryRomSerial(2, 373) = "34KEY"
    aryRomSerial(3, 373) = "99EX"
    
    aryRomSerial(0, 374) = "50F0063006B00000014341D8153"    'HH Serial 14341D8153  HH7693 34 4/14/2016 - NAJ - Westbury Lumber
    aryRomSerial(1, 374) = "SCANWEDGE"
    aryRomSerial(2, 374) = "34KEY"
    aryRomSerial(3, 374) = "99EX"
    
    aryRomSerial(0, 375) = "50F0063006B00000016140D8161"    'HH Serial 16140D8161  HH7743 34 6/10/2016 - NAJ - Krueger
    aryRomSerial(1, 375) = "SCANWEDGE"
    aryRomSerial(2, 375) = "34KEY"
    aryRomSerial(3, 375) = "99EX"
    
    aryRomSerial(0, 376) = "50F0063006B00000016149D809C"    'HH Serial 16149D809C  HH7754 34 6/22/2016 - NAJ - Weyerhauser
    aryRomSerial(1, 376) = "SCANWEDGE"
    aryRomSerial(2, 376) = "34KEY"
    aryRomSerial(3, 376) = "99EX"
    
    aryRomSerial(0, 377) = "50F0063006B00000016140D81B9"    'HH Serial 16140D81B9  HH7742 34 6/23/2016 - RM - S&J Lumber
    aryRomSerial(1, 377) = "SCANWEDGE"
    aryRomSerial(2, 377) = "34KEY"
    aryRomSerial(3, 377) = "99EX"
    
    aryRomSerial(0, 378) = "50F0063006B00000016149D809D"    'HH Serial 16149D809D  HH7753 34 6/28/2016 - NAJ - Valley Hardwoods
    aryRomSerial(1, 378) = "SCANWEDGE"
    aryRomSerial(2, 378) = "34KEY"
    aryRomSerial(3, 378) = "99EX"
    
    aryRomSerial(0, 379) = "50F0063006B00000016140D814E"    'HH Serial 16140D814E  HH7780 34 7/8/2016 - NAJ - Quality MI
    aryRomSerial(1, 379) = "SCANWEDGE"
    aryRomSerial(2, 379) = "34KEY"
    aryRomSerial(3, 379) = "99EX"
    
    aryRomSerial(0, 380) = "50F0063006B00000016140D816C"    'HH Serial 16140D816C  HH7781 34 7/8/2016 - NAJ - Quality MI
    aryRomSerial(1, 380) = "SCANWEDGE"
    aryRomSerial(2, 380) = "34KEY"
    aryRomSerial(3, 380) = "99EX"
    
    aryRomSerial(0, 381) = "50F0063006B00000016149D808F"    'HH Serial 16149D808F  HH7782 34 7/15/2016 - RM - Weyerhauser
    aryRomSerial(1, 381) = "SCANWEDGE"
    aryRomSerial(2, 381) = "34KEY"
    aryRomSerial(3, 381) = "99EX"
    
    aryRomSerial(0, 382) = "50F0063006B00000016149D8078"    'HH Serial 16149D8078  HH7793 34 7/20/2016 - NAJ - Northwood Resources
    aryRomSerial(1, 382) = "SCANWEDGE"
    aryRomSerial(2, 382) = "34KEY"
    aryRomSerial(3, 382) = "99EX"
    
    aryRomSerial(0, 383) = "50F0063006B00000016140D8215"    'HH Serial 16140D8215  HH7794 34 8/3/2016 - NAJ - Rock Lumber
    aryRomSerial(1, 383) = "SCANWEDGE"
    aryRomSerial(2, 383) = "34KEY"
    aryRomSerial(3, 383) = "99EX"
    
    aryRomSerial(0, 384) = "50F0063006B00000016140D816B"    'HH Serial 16140D819B  HH7796 34 8/5/2016 - NAJ - Rock Lumber
    aryRomSerial(1, 384) = "SCANWEDGE"
    aryRomSerial(2, 384) = "34KEY"
    aryRomSerial(3, 384) = "99EX"
    
    aryRomSerial(0, 385) = "50F0063006B00000016207D813A"    'HH Serial 16207D813A  HH7805 34 8/17/2016 - NAJ - Ram Forest Products
    aryRomSerial(1, 385) = "SCANWEDGE"
    aryRomSerial(2, 385) = "34KEY"
    aryRomSerial(3, 385) = "99EX"
    
    aryRomSerial(0, 386) = "50F0063006B00000016207D810E"    'HH Serial 16207D810E  HH7828 34 9/16/2016 - NAJ - Catskill Forest Products
    aryRomSerial(1, 386) = "SCANWEDGE"
    aryRomSerial(2, 386) = "34KEY"
    aryRomSerial(3, 386) = "99EX"
    
    aryRomSerial(0, 387) = "50F0063006B00000016208D80E8"    'HH Serial 16208D80E8  HH7830 34 9/16/2016 - NAJ - Gates Custom Milling
    aryRomSerial(1, 387) = "SCANWEDGE"
    aryRomSerial(2, 387) = "34KEY"
    aryRomSerial(3, 387) = "99EX"
    
    aryRomSerial(0, 388) = "50F0063006B00000016089D8245"    'HHD7737  Serial 16089D8245   34 9/20/2016 - Mike Johnson Demo Unit.
    aryRomSerial(1, 388) = "SCANWEDGE"
    aryRomSerial(2, 388) = "34KEY"
    aryRomSerial(3, 388) = "99EX"
    
    aryRomSerial(0, 389) = "50F0063006B00000016258D81B3"    'HHD7885  Serial 16258D81B3   34 10/17/2016 - High Quality Hardwoods - NAJ.
    aryRomSerial(1, 389) = "SCANWEDGE"
    aryRomSerial(2, 389) = "34KEY"
    aryRomSerial(3, 389) = "99EX"
    
    aryRomSerial(0, 390) = "50F0063006B00000016178D8079"    'HH7898  Serial 16178D8079   34 11/18/2016 - Stella Jones (Goshen) - RTM
    aryRomSerial(1, 390) = "SCANWEDGE"
    aryRomSerial(2, 390) = "34KEY"
    aryRomSerial(3, 390) = "99EX"
    
    aryRomSerial(0, 391) = "50F0063006B00000016257D822F"    'HH7903  Serial 16257D822F   34 11/22/2016 - RAM Forest Products - RTM
    aryRomSerial(1, 391) = "SCANWEDGE"
    aryRomSerial(2, 391) = "34KEY"
    aryRomSerial(3, 391) = "99EX"
    
    aryRomSerial(0, 392) = "50F0063006B00000016257D8218"    'HH7905  Serial 16257D8218   34 11/22/2016 - RAM Forest Products - RTM
    aryRomSerial(1, 392) = "SCANWEDGE"
    aryRomSerial(2, 392) = "34KEY"
    aryRomSerial(3, 392) = "99EX"
    
    aryRomSerial(0, 393) = "50F0063006B00000016258D8182"    'HH7908  Serial 16258D8182   34 12/01/2016 - Meador Wood Products - RTM
    aryRomSerial(1, 393) = "SCANWEDGE"
    aryRomSerial(2, 393) = "34KEY"
    aryRomSerial(3, 393) = "99EX"
    
    aryRomSerial(0, 394) = "50F0063006B0000001617BD8079"    'HH7898  Serial 1617BD8079   34 11/21/2016 - Stella-Jones Goshen Division - RTM
    aryRomSerial(1, 394) = "SCANWEDGE"
    aryRomSerial(2, 394) = "34KEY"
    aryRomSerial(3, 394) = "99EX"
    
    aryRomSerial(0, 395) = "50F0063006B00000016257D822F"    'HH7903  Serial 16257D822F   34 11/22/2016 - RAM Forest Products - RTM
    aryRomSerial(1, 395) = "SCANWEDGE"
    aryRomSerial(2, 395) = "34KEY"
    aryRomSerial(3, 395) = "99EX"
    
    aryRomSerial(0, 396) = "50F0063006B00000016257D81BB"    'HH7910  Serial 16257D81BB   34 12/08/2016 - Wilson Forest Products - RTM
    aryRomSerial(1, 396) = "SCANWEDGE"
    aryRomSerial(2, 396) = "34KEY"
    aryRomSerial(3, 396) = "99EX"
    
    aryRomSerial(0, 397) = "50F0063006B00000016261D8214"    'HH7915  Serial 16261D8214   34 12/12/2016 - Meador Wood Products - NAJ
    aryRomSerial(1, 397) = "SCANWEDGE"
    aryRomSerial(2, 397) = "34KEY"
    aryRomSerial(3, 397) = "99EX"

    aryRomSerial(0, 398) = "50F0063006B00000016258D81C6"    'HH7917  Serial 16258D81C6   34 12/13/2016 - Meador Wood Products - NAJ
    aryRomSerial(1, 398) = "SCANWEDGE"
    aryRomSerial(2, 398) = "34KEY"
    aryRomSerial(3, 398) = "99EX"

    aryRomSerial(0, 399) = "50F0063006B00000012188D0212"    'HH2302  Serial 12188D0212   34 1/24/2017 - Frick Lumber - NAJ
    aryRomSerial(1, 399) = "SCANWEDGE"
    aryRomSerial(2, 399) = "34KEY"
    aryRomSerial(3, 399) = "99EX"

    aryRomSerial(0, 400) = "50F0063006B00000016178D80F2"    'HH7941  Serial 16178D80F2   34 3/22/2017 - Dimension Hardwoods - NAJ
    aryRomSerial(1, 400) = "SCANWEDGE"
    aryRomSerial(2, 400) = "34KEY"
    aryRomSerial(3, 400) = "99EX"

    aryRomSerial(0, 401) = "50F0063006B00000016178D80E5"    'HH7943  Serial 16178D80E5   34 3/22/2017 - Dimension Hardwoods - NAJ
    aryRomSerial(1, 401) = "SCANWEDGE"
    aryRomSerial(2, 401) = "34KEY"
    aryRomSerial(3, 401) = "99EX"
    
    aryRomSerial(0, 402) = "50F0063006B00000016349D805B"    'HH7953  Serial 16349D805B   34 4/04/2017 - Consdan International LLC - NAJ
    aryRomSerial(1, 402) = "SCANWEDGE"
    aryRomSerial(2, 402) = "34KEY"
    aryRomSerial(3, 402) = "99EX"

    aryRomSerial(0, 403) = "50F0063006B00000016178D8068"    'HH7892  Serial 16178D8068   34 4/19/2017 - Edward S. Kocjancic - NAJ
    aryRomSerial(1, 403) = "SCANWEDGE"
    aryRomSerial(2, 403) = "34KEY"
    aryRomSerial(3, 403) = "99EX"

    aryRomSerial(0, 404) = "50F0063006B00000017077D8514"    'HH7959  Serial 17077D8514   55 5/08/2017 - Consdan - NAJ - Dustin Reamy
    aryRomSerial(1, 404) = "SCANWEDGE"
    aryRomSerial(2, 404) = "55KEY"
    aryRomSerial(3, 404) = "99EX"

    aryRomSerial(0, 405) = "50F0063006B00000017064D8145"    'HH7958  Serial 17064D8145   55 5/08/2017 - Consdan - NAJ - Dustin Reamy
    aryRomSerial(1, 405) = "SCANWEDGE"
    aryRomSerial(2, 405) = "55KEY"
    aryRomSerial(3, 405) = "99EX"

    aryRomSerial(0, 406) = "50F0063006B00000016178D8090"    'HH7978  Serial 16178D8090   34 6/13/2017 - Pallet One - NAJ
    aryRomSerial(1, 406) = "SCANWEDGE"
    aryRomSerial(2, 406) = "34KEY"
    aryRomSerial(3, 406) = "99EX"

    aryRomSerial(0, 407) = "50F0063006B00000016178D80B1"    'HH7980  Serial 16178D80B1   34 3/22/201 - Tucker Timber Products - NAJ
    aryRomSerial(1, 407) = "SCANWEDGE"
    aryRomSerial(2, 407) = "34KEY"
    aryRomSerial(3, 407) = "99EX"

    aryRomSerial(0, 408) = "50F0063006B00000016178D80C5"    'HH7982  Serial 16178D80C5   34 3/22/2018 - Jackson Brothers Lumber - NAJ
    aryRomSerial(1, 408) = "SCANWEDGE"
    aryRomSerial(2, 408) = "34KEY"
    aryRomSerial(3, 408) = "99EX"

    
    aryRomSerial(0, 409) = "50F0063006B00000016178D804A"    'HH7984  Serial 16178D804A   34 6/27/2017 - Legend Lumber - NAJ
    aryRomSerial(1, 409) = "SCANWEDGE"
    aryRomSerial(2, 409) = "34KEY"
    aryRomSerial(3, 409) = "99EX"

    aryRomSerial(0, 410) = "50F0063006B00000016178D80B9"    'HH7990  Serial 16178D80B9   34 7/18/2017 - Pennyrile Lumber - NAJ
    aryRomSerial(1, 410) = "SCANWEDGE"
    aryRomSerial(2, 410) = "34KEY"
    aryRomSerial(3, 410) = "99EX"
    
    aryRomSerial(0, 411) = "50F0063006B00000016178D80AF"    'HH7992  Serial 16178D80AF   34 7/25/2017 - Keweenaw Land Association - NAJ
    aryRomSerial(1, 411) = "SCANWEDGE"
    aryRomSerial(2, 411) = "34KEY"
    aryRomSerial(3, 411) = "99EX"

    aryRomSerial(0, 412) = "50F0063006B00000016178D80BB"    'HH7997  Serial 16178D80BB   34 8/22/2017 - Pennyrile Sawmill - NAJ
    aryRomSerial(1, 412) = "SCANWEDGE"
    aryRomSerial(2, 412) = "34KEY"
    aryRomSerial(3, 412) = "99EX"
 
    aryRomSerial(0, 413) = "50F0063006B00000016178D80B4"    'HH7998  Serial 16178D80B4   34 8/22/2017 - Pennyrile Sawmill - NAJ
    aryRomSerial(1, 413) = "SCANWEDGE"
    aryRomSerial(2, 413) = "34KEY"
    aryRomSerial(3, 413) = "99EX"

    aryRomSerial(0, 414) = "50F0063006B00000016178D80C5"  'HH Serial 16178D80C5 HH7982 34 Key Buskirk Lumber 10/12/2017
    aryRomSerial(1, 414) = "SCANWEDGE"
    aryRomSerial(2, 414) = "34KEY"
    aryRomSerial(3, 414) = "99EX"

    aryRomSerial(0, 415) = "50F0063006B00000016178D8047"  'HH Serial 16178D8047 HH8007 34 Key Northwood Resources 03/22/2018
    aryRomSerial(1, 415) = "SCANWEDGE"
    aryRomSerial(2, 415) = "34KEY"
    aryRomSerial(3, 415) = "99EX"

    aryRomSerial(0, 416) = "50F0063006B00000016178D806A"  'HH Serial 16178D806A HH8008 34 Key Buskirk Lumber 10/12/2017
    aryRomSerial(1, 416) = "SCANWEDGE"
    aryRomSerial(2, 416) = "34KEY"
    aryRomSerial(3, 416) = "99EX"

    aryRomSerial(0, 417) = "50F0063006B00000016178D805D"  'HH Serial 16178D805D HH8009 34 Key Tucker Timber 03/21/2018
    aryRomSerial(1, 417) = "SCANWEDGE"
    aryRomSerial(2, 417) = "34KEY"
    aryRomSerial(3, 417) = "99EX"

    aryRomSerial(0, 418) = "50F0063006B00000016178D80CD"  'HH Serial 16178D80CD HH8010 34 Key Buskirk Lumber 10/12/2017
    aryRomSerial(1, 418) = "SCANWEDGE"
    aryRomSerial(2, 418) = "34KEY"
    aryRomSerial(3, 418) = "99EX"

    aryRomSerial(0, 419) = "50F0063006B00000016178D80BD"  'HH Serial 16178D80BD HH8011 34 Key Krueger Lumber 03/22/2018
    aryRomSerial(1, 419) = "SCANWEDGE"
    aryRomSerial(2, 419) = "34KEY"
    aryRomSerial(3, 419) = "99EX"
    
    
    
    aryRomSerial(0, 420) = "50F0063006B00000017293D80DE"  'HH Serial 17293D80DE HH8028 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 420) = "SCANWEDGE"
    aryRomSerial(2, 420) = "34KEY"
    aryRomSerial(3, 420) = "99EX"

    aryRomSerial(0, 421) = "50F0063006B00000017293D80EB"  'HH Serial 17293D80EB HH8029 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 421) = "SCANWEDGE"
    aryRomSerial(2, 421) = "34KEY"
    aryRomSerial(3, 421) = "99EX"

    aryRomSerial(0, 422) = "50F0063006B00000017293D80F2"  'HH Serial 17293D80F2 HH8030 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 422) = "SCANWEDGE"
    aryRomSerial(2, 422) = "34KEY"
    aryRomSerial(3, 422) = "99EX"

    aryRomSerial(0, 423) = "50F0063006B00000017293D8112"  'HH Serial 17293D8112 HH8031 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 423) = "SCANWEDGE"
    aryRomSerial(2, 423) = "34KEY"
    aryRomSerial(3, 423) = "99EX"

    aryRomSerial(0, 424) = "50F0063006B00000017293D8110"  'HH Serial 17293D8110 HH8032 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 424) = "SCANWEDGE"
    aryRomSerial(2, 424) = "34KEY"
    aryRomSerial(3, 424) = "99EX"

    aryRomSerial(0, 425) = "50F0063006B00000017293D80F6"  'HH Serial 17293D80F6 HH8033 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 425) = "SCANWEDGE"
    aryRomSerial(2, 425) = "34KEY"
    aryRomSerial(3, 425) = "99EX"

    aryRomSerial(0, 426) = "50F0063006B00000017293D80EA"  'HH Serial 17293D80EA HH8034 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 426) = "SCANWEDGE"
    aryRomSerial(2, 426) = "34KEY"
    aryRomSerial(3, 426) = "99EX"

    aryRomSerial(0, 427) = "50F0063006B00000017293D8010"  'HH Serial 17293D8010 HH8035 34 Key Buskirk Lumber 11/15/2017 - RTM
    aryRomSerial(1, 427) = "SCANWEDGE"
    aryRomSerial(2, 427) = "34KEY"
    aryRomSerial(3, 427) = "99EX"
    
    aryRomSerial(0, 428) = "50F0063006B00000016178D806D"  'HH Serial 16178D806D HH8059 34 Key Gates Custom Milling 11/30/2017 - RDW
    aryRomSerial(1, 428) = "SCANWEDGE"
    aryRomSerial(2, 428) = "34KEY"
    aryRomSerial(3, 428) = "99EX"
    
    aryRomSerial(0, 429) = "50F0063006B00000016178D806D"  'HH Serial 16178D806D HH8012 34 Key S&J Lumber 11/30/2017 - RDW
    aryRomSerial(1, 429) = "SCANWEDGE"
    aryRomSerial(2, 429) = "34KEY"
    aryRomSerial(3, 429) = "99EX"
    
    aryRomSerial(0, 430) = "50F0063006B00000017327D80BD"  'HH Serial 17327D80BD HH8186 34 Key S&J Lumber 1/31/2018 - RDW
    aryRomSerial(1, 430) = "SCANWEDGE"
    aryRomSerial(2, 430) = "34KEY"
    aryRomSerial(3, 430) = "99EX"
    
    aryRomSerial(0, 431) = "50F0063006B00000017293D8016"  'HH Serial 17293D8016 HH8176 34 Key Lincoln County Hardwoods 2/5/2018 - RDW
    aryRomSerial(1, 431) = "SCANWEDGE"
    aryRomSerial(2, 431) = "34KEY"
    aryRomSerial(3, 431) = "99EX"
    
    aryRomSerial(0, 432) = "50F0063006B00000017293D80C4"  'HH Serial 17293D80C4 HH8060 34 Key Loggers Inc 12/19/2017 - RDW
    aryRomSerial(1, 432) = "SCANWEDGE"
    aryRomSerial(2, 432) = "34KEY"
    aryRomSerial(3, 432) = "99EX"
    
    aryRomSerial(0, 433) = "50F0063006B00000017320D8300"  'HH Serial 17320D8300 HH8198 34 Key Flowers Timber Co 2/6/2018 - RDW
    aryRomSerial(1, 433) = "SCANWEDGE"
    aryRomSerial(2, 433) = "34KEY"
    aryRomSerial(3, 433) = "99EX"
    
    aryRomSerial(0, 434) = "50F0063006B00000016299D8125"  'HH Serial 16299D8125 HH8231 34 Key Harold White 5/29/2018 - NAJ
    aryRomSerial(1, 434) = "SCANWEDGE"
    aryRomSerial(2, 434) = "34KEY"
    aryRomSerial(3, 434) = "99EX"

    aryRomSerial(0, 435) = "50F0063006B00000017088D81B2"  'HH Serial  17088D81B2  HH8230 34 Key Harold White 5/29/2018 - NAJ
    aryRomSerial(1, 435) = "SCANWEDGE"
    aryRomSerial(2, 435) = "34KEY"
    aryRomSerial(3, 435) = "99EX"

    aryRomSerial(0, 436) = "50F0063006B00000017088D8145"  'HH Serial 17088D81B2  HH8229 34 Key Harold White 5/29/2018 - NAJ
    aryRomSerial(1, 436) = "SCANWEDGE"
    aryRomSerial(2, 436) = "34KEY"
    aryRomSerial(3, 436) = "99EX"
    
    
    aryRomSerial(0, 437) = "50F0063006B00000018165D8205"  'HH Serial  11/30/2018 18165D8205 HH8338 34 Key Wilson Forest Products
    aryRomSerial(1, 437) = "SCANWEDGE"
    aryRomSerial(2, 437) = "34KEY"
    aryRomSerial(3, 437) = "99EX"
   
        
    #If APPFORGE Then
        gstrHHPSerialNum = GetHHSerialAdjustment(AFCoreLib.ROMSerialNumber)
    #Else
'        gstrHHPSerialNum = GetHHSerialAdjustment("50F0063006B000000")
    #End If
    
    
    
    fLicense = False
 '       MsgBox "This Handheld is Not Licensed for the Mobile Limbs System.  Contact eLimbs at 888.520.1951 for support and licnesing (" & gstrHHPSerialNum & ")"
    
    KeyF1 = 112
    KeyF2 = 113
    KeyF3 = 114
    KeyF3_99EX = 230
    KeyF4 = 115
    KeyF5 = 116
    KeyShift = 16
    
    KeyF6 = 96
    
    KeyF1_Keyboard = 112
    KeyF2_Keyboard = 113
    KeyF3_Keyboard = 114
    KeyF4_Keyboard = 115
    KeyF5_Keyboard = 116
    KeyF6_Keyboard = 117
    KeyF7_Keyboard = 118
    KeyF8_Keyboard = 119
    
    KeyBackSlash = 220
    keyEnter = 13
    KeyScanButton = 42
    
    If gSettings.DisableScanButton = "YES" Then KeyScanButton = 2000
    

    fLicense = False
    
    For I = 0 To UBound(aryRomSerial, 2)
        If UCase(aryRomSerial(0, I)) = UCase(gstrHHPSerialNum) Then
            'Continue
            fLicense = True
            gstrScanner = "SCANWEDGE"
            gstrHHModel = Trim(UCase(aryRomSerial(3, I)))
            If sc(aryRomSerial(4, I), "LOGTALLYONLY-RESTRICTED") = True Then
                gfRestrictNonHandheldFunctions = True
            Else
                gfRestrictNonHandheldFunctions = False
            End If
            
''            If I = 172 Then
''                MsgBox aryRomSerial(3, I) & " - " & aryRomSerial(2, I) & " - " & aryRomSerial(0, I)
''            End If
''
            If aryRomSerial(3, I) = "9900" And aryRomSerial(2, I) = "35KEY" Then
                KeyHelp = 227
                KeyF1 = 227
                KeyF2 = 228
                KeyF3 = 230
                KeyF4 = 233
                KeyF5 = 231
                KeyF6 = 232
                gstrHHKeyPad = aryRomSerial(2, I)
            ElseIf aryRomSerial(3, I) = "99EX" And aryRomSerial(2, I) = "34KEY" Then
                KeyHelp = 227
                KeyF1 = 227
                KeyF2 = 228
                KeyF3 = 230
                KeyF4 = 233
                KeyF5 = 234
                KeyF6 = 235
                KeyF7 = 236
                KeyF8 = 237
                KeyBackSlash = 220
                gstrHHKeyPad = aryRomSerial(2, I)
           
            ElseIf aryRomSerial(3, I) = "99EX" And aryRomSerial(2, I) = "55KEY" Then
                KeyHelp = 227
                KeyF1 = 227
                KeyF2 = 228
                KeyF3 = 230
                KeyF4 = 233
                KeyF5 = 234
                KeyF6 = 235
                KeyF7 = 236
                KeyF8 = 237
                KeyBackSlash = 191
                gstrHHKeyPad = aryRomSerial(2, I)
                
            ElseIf aryRomSerial(3, I) = "9900" And aryRomSerial(2, I) = "56KEY" Then
                KeyHelp = KeyHelp2
                KeyF1 = 227
                KeyF2 = 228
                KeyF3 = 230
                KeyF4 = 233
                KeyF5 = 231
                KeyF6 = 232
                gstrHHKeyPad = aryRomSerial(2, I)

            Else
                KeyHelp = 112
            End If
            Exit For
        End If
    Next
    
    If IsNumeric(gSettings.KeyF1Alt) = True Then
        KeyHelp = CInt(gSettings.KeyF1Alt)
    End If
        
    
'fLicense = True
'fLicense = True

    
On Error GoTo ErrorHandler
'GoTo goto8
    Set FileMgr = New CFileManager

    intError = 0
        Call FileMgr.OpenDirectory(App.Path & "\db", True)

'''intError = 10
'''
'''    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\tblFootage.pdb", afFileModeOpen)
'''    Set TempFile = Nothing
'''Goto10:
intError = 1

    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\tblCorrections.pdb", afFileModeOpen)
    Set TempFile = Nothing
Goto1:
intError = 2

    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\tblCMBF.pdb", afFileModeOpen)
    Set TempFile = Nothing
Goto2:
intError = 3

    
        
        
goto3:
intError = 4
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\tblLogUtility.pdb", afFileModeOpen)
    Set TempFile = Nothing
goto4:
intError = 5
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\tblmoneyhistory.pdb", afFileModeOpen)
    Set TempFile = Nothing
goto5:
intError = 6
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\Utility.pdb", afFileModeOpen)
    Set TempFile = Nothing
    

    
goto6:

intError = 7
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\tblTally.pdb", afFileModeOpen)
    Set TempFile = Nothing

goto7:

intError = 8
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\LogScale.pdb", afFileModeOpen)
    Set TempFile = Nothing
  
Goto8:
intError = 9
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\Rec.pdb", afFileModeOpen)
    Set TempFile = Nothing
    
Goto9:
intError = 10
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\LogPurch.pdb", afFileModeOpen)
    Set TempFile = Nothing

Goto10:
intError = 10.5
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\LogInv.pdb", afFileModeOpen)
    Set TempFile = Nothing

Goto10_5:
intError = 11
    Set TempFile = FileMgr.OpenAsBinary(App.Path & "\db\Lookup.pdb", afFileModeOpen)
    Set TempFile = Nothing
    
Goto11:
intError = 12

    OpentblLogsDatabase
    If dbtblLogs = 0 Or PDBGetNumFields(dbtblLogs) = 0 Then
        MsgBox "There is A Critical Error With the Log DATABASE, Continuing Will Certainly Result in Loss of All Tally Data. Contact Technical Support Immediately!"
        'End
    End If
    ClosetblLogsDatabase
    dbtblLogs = 0
    
    
    If fLicense = False Then
        MsgBox "mLOGS v" & App.Major & "." & App.Minor & "." & App.Revision & " - This Handheld is Not Licensed for the Mobile Limbs System.  Contact eLimbs at 888.520.1951 for support and licnesing (" & gstrHHPSerialNum & ")"
        End
        Exit Sub
    End If
    closealldatabases
       
    lblVersion.Caption = "(v" & App.Major & "." & App.Minor & "." & App.Revision & ") "
    fLoad = True
    Call UpdateUserList(cbo, True)

    Exit Sub

ErrorHandler:
    If intError = 0 Then
    
    ElseIf intError = 1 Then
        CreateDatabasePDB dbtblCorrections, App.Path & "\db\" & "tblcorrections", tblCorrections_Schema
        ClosetblCorrectionsDatabase

        GoTo Goto1
    ElseIf intError = 2 Then
        CreateDatabasePDB dbtblCMBF, App.Path & "\db\" & "tblCMBF", tblCMBF_Schema
        ClosetblCMBFDatabase
        GoTo Goto2
    
    ElseIf intError = 3 Then
        MsgBox "Contact Technical Support at 888.520.1951 or Support@eLIMBS.com.  There is a Critical Error in your LOGS PDB FILE!"
        End
        
'    ElseIf intError = 5 Then
'        CreateDatabasePDB dbtblMoneyHistory, App.Path & "\db\" & "tblmoneyhistory", tblMoneyHistory_Schema
'        ClosetblMoneyHistoryDatabase
'        GoTo goto5
    ElseIf intError = 6 Then
        CreateDatabasePDB dbUtility, App.Path & "\db\" & "Utility", Utility_Schema
        CloseUtilityDatabase
        GoTo goto6
    ElseIf intError = 7 Then
        CreateDatabasePDB dbtblTally, App.Path & "\db\" & "tblTally", tblTally_Schema
        ClosetblTallyDatabase
        GoTo goto7
    ElseIf intError = 8 Then
        CreateDatabasePDB dbLogScale, App.Path & "\db\" & "LogScale", LogScale_Schema
        CloseLogScaleDatabase
        GoTo Goto8
    ElseIf intError = 9 Then
        CreateDatabasePDB dbRec, App.Path & "\db\" & "Rec", Rec_Schema
        CloseRecDatabase
        GoTo Goto9
'''  ElseIf intError = 10 Then
'''        CreateDatabasePDB dbtblFootage, App.Path & "\db\" & "tblFootage", tblFootage_Schema
'''        ClosetblFootageDatabase
'''        GoTo Goto10
        
  ElseIf intError = 10 Then
        CreateDatabasePDB dbLogPurch, App.Path & "\db\" & "LogPurch", LogPurch_Schema
        CloseLogPurchDatabase
        GoTo Goto10
  ElseIf intError = 10.5 Then
        CreateDatabasePDB dbLogInv, App.Path & "\db\" & "LogInv", LogInv_Schema
        CloseLogInvDatabase
        
        GoTo Goto10_5
  ElseIf intError = 11 Then
        CreateDatabasePDB dbLookup, App.Path & "\db\" & "Lookup", Lookup_Schema
        CloseLookupDatabase
        
        
        GoTo Goto11

'''  ElseIf intError = 12 Then
'''        CreateDatabasePDB dbtblLookupHeader, App.Path & "\db\" & "tblLookupHeader", tblLookupHeader_Schema
'''        ClosetblLookupHeaderDatabase
'''       GoTo goto12
'''  ElseIf intError = 13 Then
'''        CreateDatabasePDB dbtblLookupDetail, App.Path & "\db\" & "tblLookupDetail", tblLookupDetail_Schema
'''        ClosetblLookupDetailDatabase
'''       GoTo goto13
  
  ElseIf intError = 4 Then
        Set TempFile = Nothing
        Set FileMgr = Nothing

        If MsgBox("System Databases are Missing, Would you like to Restore the Last Backup Now?", vbYesNo) = vbYes Then
            Call BackuptoSD(lblBackup, True, "", True)
            fRestore = True
            Call RestoreBackup
        End If
        MsgBox "Restore Completed Log System Will close you will need to reopen mLOGS when completed!"
        End
        GoTo Goto9
    End If

End Sub


Private Sub mnuAbout_Click()

End Sub

Private Sub txtUserLoginNameHH_GotFocus()
    txtUserLoginNameHH.BackColor = &H80FFFF
    Call Got_Focus(txtUserLoginNameHH)

End Sub

Private Sub txtUserLoginNameHH_LostFocus()
    If UCase(txtUserLoginNameHH.text) = "KEY" Then
        'then do nothing...let it be
    Else
        txtUserLoginNameHH.BackColor = vbWhite

        txtUserLoginNameHH.BackColor = vbWhite
        Call AppForge_Split(GetUserData(Trim(UCase(txtUserLoginNameHH.text)), "USERNAME"), pstrSplit, ",")

        lblUserLoginNameHH.tag = pstrSplit.ary(0)
        lblUserLoginNameHH.Caption = pstrSplit.ary(1)
        txtUserLoginNameHH.text = pstrSplit.ary(1)
    End If

End Sub

Private Sub txtUserPasswordHH_GotFocus()
    txtUserPasswordHH.BackColor = &H80FFFF
    Call Got_Focus(txtUserPasswordHH)
End Sub

Private Sub txtUserPasswordHH_LostFocus()
    txtUserPasswordHH.BackColor = vbWhite
End Sub

Private Sub tmrBackup_Timer()
    gintTimer = gintTimer + 1

    If gintTimer > 5 Then
        gintTimer = 0
        Call BackuptoSD(lblBackup)              ' Should this be at 5 minutes Leith.Stetson 07.07.2006
    End If

End Sub

