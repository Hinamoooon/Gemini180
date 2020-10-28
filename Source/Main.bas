Attribute VB_Name = "Main"
Option Explicit '�ϐ��̐錾�K�{

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
Dim DA As String    'Device Primary address�w��p
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

SlitStepsPerUnit = Range("D3").value        '7mm�̃X���b�g��145steps/unit

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
Call ibfind(BOARD, UDB) '�f�o�C�X�����s���ȂƂ��Ɏg�p����Ă��Ȃ��f�o�C�X��
If UDB < 0 Then         '�I�[�v�����ď��������s��(ibfind�̐���(?)���삳��̃v���O�������)
    MsgBox "Board couldn't be found."  '�G���[����UDB��"-1"��Ԃ��B(Refer to "NI-488.2 reference manual for windows")
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

'�ݒ荀��
EOSV = &HD          '�I�[����(EOS: End of Scripts)��"0D" (=<CR>)�B���ʃo�C�g�B
V = EOSV + &H1400   'EOS��ʃo�C�g��A��C�����K�p�B(Refer to "NI-488.2 reference manual for windows")
Call ibeos(UDG, V)   '��2�s�ɏ]��EOS�ݒ�'        C�͕���bit�Ȃ��Ƃ������ƁH
Call ibtmo(UDG, T300ms)  '�^�C���A�E�g���Ԃ̐ݒ�'

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
ICH = " "                   'Gemini180 controller�̏ꏊ�ɉ�����
Call ibwrt(UDG, ICH)        '"B"(for BOOT) or "F"(for MAIN)��Ԃ��B
If (ibsta And EERR) Then    'Error code. Refer "NIGLOBAL.bas" for EERR. EERR=&H8000 (16��bit��ON�Ȃ�G���[�j
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
TEMP = MsgBox("Do you need to Initialize Motors ?", vbYesNo + vbQuestion, "�m�F")
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
If ibsta% And CMPL Then Comment1 = " CMPL"  'means I/O completed (����j
If ibsta% And LOK Then Comment1 = " LOK"
If ibsta% And RREM Then Comment1 = " REM"
If ibsta% And CIC Then Comment1 = " CIC"
If ibsta% And AATN Then Comment1 = " ATN"
If ibsta% And TACS Then Comment1 = " TACS"
If ibsta% And LACS Then Comment1 = " LACS"
If ibsta% And DTAS Then Comment1 = " DTAS"
If ibsta% And DCAS Then Comment1 = " DCAS"

'Print "iberr= "; iberr%;
If iberr% = EDVR Then Comment2 = " EDVR <DOS Error>"    'ibsta <> EERR �̂Ƃ��� EDVR (iberr value=0)������H
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
PBUF = Left(RBUF, InStr(RBUF, CR) - 1)  'CR�̍����݂̂ɐؒf
PBUF = Right(PBUF, Len(PBUF) - 1)       '"o"���폜
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
    Call ibrd(UDG, RBUF) 'Recieve "o". (�s�v�H)
    
    'Wait while Gemini180 is moving     ' "E"����Ȃ��A"Z453"�ł́H
    Do
        Range("H4").value = ""
        Application.Wait [Now() + "0:00:00.5"]
        Call READ_WORKING_ABS_POSITION
        Application.StatusBar = "Motor is moving..."
        ICH = "E"
        Call ibwrt(UDG, ICH)
        Call ibrd(UDG, RBUF)             '"o"�ɑ����āA���[�^�[��busy�Ȃ�"q"�Anot busy�Ȃ�"z"��Ԃ�
        PBUF = Left(RBUF, InStr(RBUF, "o") + 1)   'RBUF����o�̈ʒu����肵�Ao�̎��̕����܂ł�PBUF��
    Loop While PBUF = "oq"   'busy�Ȃ珈�����J��Ԃ�
    
    'Over write position
    Call READ_WORKING_ABS_POSITION
    
    'Reset status ber
    Application.StatusBar = False
End If

End Sub


Sub IncreaseWavelength()
'K4�ɓ��͂��ꂽ�l��������

Dim INC As Double
INC = Range("K4").value

'K4�̒l�Ɋ�Â��AH4���X�V
Range("H4").value = Range("H4").value + INC

Call MOVE_WORKING_ABS_POSITION

End Sub


Sub MotorInit() '�ꎞ�I��ibtmo(UDG, T100s)�ɂ��ׂ� (refer Spectrometer control)

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

MsgBox "Wait 50 sec!"    '30�b�҂��Ȃ��ƃG���[���o��'
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
PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'

MsgBox "Slit Speed: " & PBUF, 64, "Slit Read Speed"

End Sub


Sub SlitReadPosition()

Dim NUM As Integer

For NUM = 0 To 2
    Slit = Cells(8 + NUM, 13).value
    
    OBUF = "j0," & Str(Slit) & Chr(13) '���݂̃X���b�g�ʒu��ǂ�'
    Call ibwrt(UDG, OBUF)
    Call ibrd(UDG, RBUF)
    
    PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
    PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
    
    WSlitPOS = PBUF / SlitStepsPerUnit
    Cells(8 + NUM, 8).value = WSlitPOS
Next

End Sub


Sub SlitSetPositionBox() '�X���b�g�ԍ����ʂɎw�肵�ăX���b�g�𓮂���'

Slit = InputBox("�X���b�g�ԍ��̓���" & Chr(13) & "0, 2, 3�̐����œ��́B", Title:="Slit Set Position") '�X���b�g1�͖���'

Do While (SlitPOS > 1120 Or SlitPOS < 0)
WSlitPOS = InputBox("�X���b�g��(mm)�̓��́B" & Chr(13) _
& "0(mm)����2(mm)�܂ŉ\�B", Title:="Set Slit Position")
SlitPOS = WSlitPOS * SlitStepsPerUnit
Loop

OBUF = "j0," & Str(Slit) & Chr(13)  'Slit read position
Call ibwrt(UDG, OBUF)
Call ibrd(UDG, RBUF)
PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'

SlitPOS = SlitPOS - CInt(PBUF)

OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13) 'Slit move rerative
Call ibwrt(UDG, OBUF)
Call ibrd(UDG, RBUF)

End Sub

Sub SlitSetPosition() '�Z������l��ǂݎ���ăX���b�g3�𓮂���'

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
        
        OBUF = "j0," & Str(Slit) & Chr(13) '���݂̃X���b�g�ʒu��ǂ�'
        Call ibwrt(UDG, OBUF)
        Call ibrd(UDG, RBUF)
        
        PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
        PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
        
        SlitPOS = SlitPOS - CInt(PBUF) '���ΓI�ɂ����瓮�����΂������v�Z'
        
        OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13) '�X�e�b�v�œ�����'
        Call ibwrt(UDG, OBUF)
        Call ibrd(UDG, RBUF)
    End If

    'Wait while Gemini180 is moving     ' "E"����Ȃ��A"Z453"�ł́H
    Do
        Cells(8 + NUM, 8).value = ""
        Application.Wait [Now() + "0:00:00.5"]
        Application.StatusBar = "Motor is moving..."
        ICH = "E"
        Call ibwrt(UDG, ICH)
        Call ibrd(UDG, RBUF)             '"o"�ɑ����āA���[�^�[��busy�Ȃ�"q"�Anot busy�Ȃ�"z"��Ԃ�
        PBUF = Left(RBUF, InStr(RBUF, "o") + 1)   'RBUF����o�̈ʒu����肵�Ao�̎��̕����܂ł�PBUF��
    Loop While PBUF = "oq"   'busy�Ȃ珈�����J��Ԃ�
    Application.StatusBar = False
Next

If BreakFlag = False Then
    Call SlitReadPosition
End If

End Sub




'#####################################
'TRIAX���̃R�}���h�͂ǂ�H
'
'Motor busy check��TRIAX�R�}���h�ł��ׂ��H
'
'#####################################
