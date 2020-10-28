Attribute VB_Name = "Main"
Option Explicit '変数の宣言必須

'Public Variable Declaration
'##############################################
Public BOARD As String
Public DEVICE As String
Public UDB As Integer  'Unit descriptor of BOARD (GPIB-USB converter)
Public UDG As Integer  'Unit descriptor of Gemini180

Public MonoStepsPerUnit As Double
Public WMONOPOS As Double
Public MONOPOS As Double

Public Slit As Integer
Public SlitPOS As Double
Public WSlitPOS As Double
Public SlitStepsPerUnit As Double
Public Freq As Integer

Public RBUF As String
Public PBUF As String
Public OBUF As String
Public Junk As String  'Buffer
Public ICH As String   'ICH$ and ACK$ are 1 byte program variables
Public ACK As String
Public CR As String
'##############################################


Sub Initialize()

'Procedure level variable Declarations
Dim DA As String    'Device Primary address指定用
Dim EOSV As Integer 'for EOS low bit setting
Dim V As Integer    'for EOS setting
Dim WaitCount As Integer   'Prevent infinite loop
Dim TEMP As Integer 'for YesNo MsgBox

'Preparation
Worksheets(1).Activate

'Setting
BOARD = "GPIB0" 'default value of USB-GPIB converter
DA = Range("D2").value
DEVICE = "DEV" & DA

SlitStepsPerUnit = Range("D3").value        '7mmのスリットで145steps/unit

RBUF = Space(132)
OBUF = Space(132)
Junk = Space(132)
ICH = Space(1)
ACK = Space(1)
CR = Chr(13)

'Start Initialization
Range("A2") = Range("A2").value & "Begin IEEE-488 Communications Setup" & vbCrLf

'Initialization of board
'ibfind: Open and initialize an interface or a user-configured device descriptor.
Call ibfind(BOARD, UDB) 'デバイス名が不明なときに使用されていないデバイスを
If UDB < 0 Then         'オープンして初期化を行う(ibfindの説明(?)西野さんのプログラムより)
    MsgBox "Board couldn't be found."  'エラー時はUDBに"-1"を返す。(Refer to "NI-488.2 reference manual for windows")
    Call ErrorFinder
    Exit Sub
End If
Range("D6").value = UDB   '@@@@@

'Initialization of Gemini180
Call ibfind(DEVICE, UDG) 'Gemini180 (Default UD of Gemini180 is 1.)'
If UDG < 0 Then
    MsgBox "Monochrometer couldn't be found."
    Call ErrorFinder
    Exit Sub
End If
Range("D7").value = UDG   '@@@@@

'設定項目
EOSV = &HD          '終端文字(EOS: End of Scripts)は"0D" (=<CR>)。下位バイト。
V = EOSV + &H1400   'EOS上位バイトはAとC両方適用。(Refer to "NI-488.2 reference manual for windows")
Call ibeos(UDG, V)   '上2行に従いEOS設定'        Cは符号bitなしということ？
Call ibtmo(UDG, T300ms)  'タイムアウト時間の設定'

'Flush anything in input buffer.
Junk = Space(132)
Call ibrd(UDG, Junk)

'Re-boot Gemini180
ICH = Chr(222)      'Value "222" force a re-boot if hung from an incomplete command.
Call ibwrt(UDG, ICH)    'You can force re-boot by sending "248" before "222" if you need.

Application.Wait (Now() + TimeValue("0:00:01"))

'Flush anything in input buffer again.
Call ibrd(UDG, Junk)

'Confirm which internal spectrometer control program we are talking to, Boot or Main.
'Send <space> 'This is WHERE AM I command.
ICH = " "                   'Gemini180 controllerの場所に応じて
Call ibwrt(UDG, ICH)        '"B"(for BOOT) or "F"(for MAIN)を返す。
If (ibsta And EERR) Then    'Error code. Refer "NIGLOBAL.bas" for EERR. EERR=&H8000 (16番bitがONならエラー）
    Call ErrorFinder
    Exit Sub
End If

Call ibrd(UDG, ICH)         'Read into "ICH"
If (ibsta And EERR) Then
    Call ErrorFinder
    Exit Sub
End If

'Display "Received" when we could read any signals from Gemini180.
Range("A2") = Range("A2").value & "Received" & vbCrLf

If ICH = "B" Then           'You have to transfer to the spectrometer controller's Main program.
    
    OBUF = "O2000" & Chr(0) 'Chr(0) = <null> (This order the spectrometer to start Main program.)
    Call ibwrt(UDG, OBUF)
    If (ibsta% And EERR) Then
        Call ErrorFinder
        Exit Sub
    End If
    Application.Wait (Now() + TimeValue("0:00:01"))
    WaitCount = 0
    
WaitReturn:

    Call ibrd(UDG, ICH)
    If (ibsta And EERR) Then
        Call ErrorFinder
        Exit Sub
    End If
    
    'Wait for Main program response. (Go to Label: WaitReturn)
    If ICH <> "*" Then
        Application.Wait (Now() + "0:00:01")
        WaitCount = WaitCount + 1
        If WaitCount < 10 Then GoTo WaitReturn
    End If

End If

'Message display
Range("A2") = Range("A2").value & "IEEE-488 Communications Established!" & vbCrLf

'Flash version check
Call FlashVersionCheck
    
'Start initialization of motors if it's needed
TEMP = MsgBox("Do you need to Initialize Motors ?", vbYesNo + vbQuestion, "確認")
If TEMP = vbYes Then
    Range("A2") = Range("A2").value & "Begin motor initializations" & vbCrLf
    Call MotorInit
End If

Range("A2") = Range("A2").value & "End..." & vbCrLf

End Sub


Sub ErrorFinder()

Dim Title As String
Dim Comment1 As String
Dim Comment2 As String

Title = "IEEE-488 ERROR"
Comment1 = "None"
Comment2 = "None"

'Print "ibsta= &H"; Hex$(ibsta%); " <";
If ibsta% And EERR Then Comment1 = " ERR"
If ibsta% And TIMO Then Comment1 = " TIMO"
If ibsta% And EEND Then Comment1 = " END"
If ibsta% And SRQI Then Comment1 = " SRQI"
If ibsta% And RQS Then Comment1 = " RQS"
If ibsta% And CMPL Then Comment1 = " CMPL"  'means I/O completed (正常）
If ibsta% And LOK Then Comment1 = " LOK"
If ibsta% And RREM Then Comment1 = " REM"
If ibsta% And CIC Then Comment1 = " CIC"
If ibsta% And AATN Then Comment1 = " ATN"
If ibsta% And TACS Then Comment1 = " TACS"
If ibsta% And LACS Then Comment1 = " LACS"
If ibsta% And DTAS Then Comment1 = " DTAS"
If ibsta% And DCAS Then Comment1 = " DCAS"

'Print "iberr= "; iberr%;
If iberr% = EDVR Then Comment2 = " EDVR <DOS Error>"    'ibsta <> EERR のときは EDVR (iberr value=0)が正常？
If iberr% = ECIC Then Comment2 = " ECIC <Not CIC>"
If iberr% = ENOL Then Comment2 = " ENOL <No Listener>"
If iberr% = EADR Then Comment2 = " EADR <Address error>"
If iberr% = EARG Then Comment2 = " EARG <Invalid argument>"
If iberr% = ESAC Then Comment2 = " ESAC <Not Sys Ctrlr>"
If iberr% = EABO Then Comment2 = " EABO <Op. aborted>"
If iberr% = ENEB Then Comment2 = " ENEB <No GPIB board>"
If iberr% = EOIP Then Comment2 = " EOIP <Async I/O in prg>"
If iberr% = ECAP Then Comment2 = " ECAP <No capability>"
If iberr% = EFSO Then Comment2 = " EFSO <File sys. error>"
If iberr% = EBUS Then Comment2 = " EBUS <Command error>"
If iberr% = ESTB Then Comment2 = " ESTB <Status byte lost>"
If iberr% = ESRQ Then Comment2 = " ESRQ <SRQ stuck on>"
If iberr% = ETAB Then Comment2 = " ETAB <Table Overflow>"

'Print "ibcnt = "; ibcnt%

Range("A2") = Range("A2").value & "ibsta => " & Comment1 & vbCrLf
Range("A2") = Range("A2").value & "iberr => " & Comment2 & vbCrLf
Range("A2") = Range("A2").value & "ibcnt => " & ibcnt & vbCrLf
Range("A2") = Range("A2").value & "End..." & vbCrLf

MsgBox Title

End Sub


Sub READ_WORKING_ABS_POSITION()

ICH = "Z62,1" & CR
Call ibwrt(UDG, ICH)
Call ibrd(UDG, RBUF)
PBUF = Left(RBUF, InStr(RBUF, CR) - 1)  'CRの左側のみに切断
PBUF = Right(PBUF, Len(PBUF) - 1)       '"o"を削除
Range("H4").value = PBUF

End Sub


Sub MOVE_WORKING_ABS_POSITION()

WMONOPOS = Range("H4")

If WMONOPOS > 1400 Or WMONOPOS < -200 Then
    MsgBox "Out of Range"
Else
    'Display status bar
    Application.StatusBar = "Processing"
    
    'Move position
    ICH = "Z61,1," & WMONOPOS & CR  'Moves wavelength of TRIAX drive according to BASE grating (1200gr/mm)
    Call ibwrt(UDG, ICH)
    Call ibrd(UDG, RBUF) 'Recieve "o". (不要？)
    
    'Wait while Gemini180 is moving     ' "E"じゃなく、"Z453"では？
    Do
        Range("H4").value = ""
        Application.Wait [Now() + "0:00:00.5"]
        Call READ_WORKING_ABS_POSITION
        Application.StatusBar = "Motor is moving..."
        ICH = "E"
        Call ibwrt(UDG, ICH)
        Call ibrd(UDG, RBUF)             '"o"に続いて、モーターがbusyなら"q"、not busyなら"z"を返す
        PBUF = Left(RBUF, InStr(RBUF, "o") + 1)   'RBUFからoの位置を特定し、oの次の文字までをPBUFに
    Loop While PBUF = "oq"   'busyなら処理を繰り返す
    
    'Over write position
    Call READ_WORKING_ABS_POSITION
    
    'Reset status ber
    Application.StatusBar = False
End If

End Sub


Sub IncreaseWavelength()
'K4に入力された値分動かす

Dim INC As Double
INC = Range("K4").value

'K4の値に基づき、H4を更新
Range("H4").value = Range("H4").value + INC

Call MOVE_WORKING_ABS_POSITION

End Sub


Sub MotorInit() '一時的にibtmo(UDG, T100s)にすべき (refer Spectrometer control)

Dim T10 As Date
Dim T20 As Date
Dim T30 As Date
Dim T40 As Date
Dim T50 As Date

T10 = Now + TimeValue("0:00:10")
T20 = T10 + TimeValue("0:00:10")
T30 = T20 + TimeValue("0:00:10")
T40 = T30 + TimeValue("0:00:10")
T50 = T40 + TimeValue("0:00:10")

Call ibwrt(UDG, "A")

MsgBox "Wait 50 sec!"    '30秒待たないとエラーが出る'
Application.Wait T10
Application.StatusBar = "Wait (20%)"
Application.Wait T20
Application.StatusBar = "Wait (40%)"
Application.Wait T30
Application.StatusBar = "Wait (60%)"
Application.Wait T40
Application.StatusBar = "Wait (80%)"
Application.Wait T50

Call READ_WORKING_ABS_POSITION
Call SlitReadPosition

Application.StatusBar = False
MsgBox "OK"

End Sub


Sub FlashVersionCheck()

Range("A2") = Range("A2").value & "Flash version check..." & vbCrLf

ICH = "y"
Call ibwrt(UDG, ICH)
Call ibrd(UDG, RBUF)
PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
PBUF = Right(PBUF, Len(PBUF) - 1)

Range("A2") = Range("A2").value & "Boot Version: " & PBUF & vbCrLf

ICH = "z"
Call ibwrt(UDG, ICH)
Call ibrd(UDG, RBUF)
PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
PBUF = Right(PBUF, Len(PBUF) - 1)

Range("A2") = Range("A2").value & "Main Version: " & PBUF & vbCrLf

End Sub


Sub EndCommunication()

Dim TEMP As Integer

'Set Grating and Slit position for 0 if it's needed
TEMP = MsgBox("Do you reset the Grating and Slit position ?", vbYesNo + vbQuestion, "Confirmation")
If TEMP = vbYes Then
    Range("A2") = Range("A2").value & "Slit position : 0 mm" & vbCrLf
    Range("H8:H10").value = 0
    Call SlitSetPosition
    Range("A2") = Range("A2").value & "Grating position : 0 nm" & vbCrLf
    Range("H4").value = 0
    Call MOVE_WORKING_ABS_POSITION
End If

Call ibonl(UDG, 0)
MsgBox "IEEE-488 Communication has been disabled."
Range("A2") = Range("A2").value & "IEEE-488 disconnected" & vbCrLf

End Sub

'--------------------------------
'Slit control
'--------------------------------

Sub SlitSetSpeed()

Slit = InputBox("Slit ?", Title:="Slit Set Speed")
Freq = InputBox("Freqency ?", Title:="Slit Set Speed")
OBUF = "g0," & Str(Slit) & "," & Str(Freq) & Chr(13)
Call ibwrt(UDG, OBUF)

End Sub


Sub SlitReadSpeed()

Slit = InputBox("Slit ?", Title:="Slit Set Speed")
OBUF = "h0," & Str(Slit) & Chr(13)  'Read position in spteps
Call ibwrt(UDG, OBUF)
Call ibrd(UDG, RBUF)
PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUFの一番左の文字を消す'

MsgBox "Slit Speed: " & PBUF, 64, "Slit Read Speed"

End Sub


Sub SlitReadPosition()

Dim NUM As Integer

For NUM = 0 To 2
    Slit = Cells(8 + NUM, 13).value
    
    OBUF = "j0," & Str(Slit) & Chr(13) '現在のスリット位置を読む'
    Call ibwrt(UDG, OBUF)
    Call ibrd(UDG, RBUF)
    
    PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
    PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUFの一番左の文字を消す'
    
    WSlitPOS = PBUF / SlitStepsPerUnit
    Cells(8 + NUM, 8).value = WSlitPOS
Next

End Sub


Sub SlitSetPositionBox() 'スリット番号を個別に指定してスリットを動かす'

Slit = InputBox("スリット番号の入力" & Chr(13) & "0, 2, 3の数字で入力。", Title:="Slit Set Position") 'スリット1は無い'

Do While (SlitPOS > 1120 Or SlitPOS < 0)
WSlitPOS = InputBox("スリット幅(mm)の入力。" & Chr(13) _
& "0(mm)から2(mm)まで可能。", Title:="Set Slit Position")
SlitPOS = WSlitPOS * SlitStepsPerUnit
Loop

OBUF = "j0," & Str(Slit) & Chr(13)  'Slit read position
Call ibwrt(UDG, OBUF)
Call ibrd(UDG, RBUF)
PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUFの一番左の文字を消す'

SlitPOS = SlitPOS - CInt(PBUF)

OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13) 'Slit move rerative
Call ibwrt(UDG, OBUF)
Call ibrd(UDG, RBUF)

End Sub

Sub SlitSetPosition() 'セルから値を読み取ってスリット3つを動かす'

Dim NUM As Integer
Dim BreakFlag As Boolean

For NUM = 0 To 2
    BreakFlag = False   'initialize
    WSlitPOS = Cells(8 + NUM, 8).value
    If WSlitPOS > 7.24 Or WSlitPOS < 0 Then
        MsgBox "Input the value from 0 to 7.24 mm."
        BreakFlag = True
        Exit For
    Else
        Slit = Cells(8 + NUM, 13).value
        SlitPOS = WSlitPOS * SlitStepsPerUnit
        
        OBUF = "j0," & Str(Slit) & Chr(13) '現在のスリット位置を読む'
        Call ibwrt(UDG, OBUF)
        Call ibrd(UDG, RBUF)
        
        PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
        PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUFの一番左の文字を消す'
        
        SlitPOS = SlitPOS - CInt(PBUF) '相対的にいくら動かせばいいか計算'
        
        OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13) 'ステップで動かす'
        Call ibwrt(UDG, OBUF)
        Call ibrd(UDG, RBUF)
    End If

    'Wait while Gemini180 is moving     ' "E"じゃなく、"Z453"では？
    Do
        Cells(8 + NUM, 8).value = ""
        Application.Wait [Now() + "0:00:00.5"]
        Application.StatusBar = "Motor is moving..."
        ICH = "E"
        Call ibwrt(UDG, ICH)
        Call ibrd(UDG, RBUF)             '"o"に続いて、モーターがbusyなら"q"、not busyなら"z"を返す
        PBUF = Left(RBUF, InStr(RBUF, "o") + 1)   'RBUFからoの位置を特定し、oの次の文字までをPBUFに
    Loop While PBUF = "oq"   'busyなら処理を繰り返す
    Application.StatusBar = False
Next

If BreakFlag = False Then
    Call SlitReadPosition
End If

End Sub




'#####################################
'TRIAX専門のコマンドはどれ？
'
'Motor busy checkはTRIAXコマンドでやるべき？
'
'#####################################
