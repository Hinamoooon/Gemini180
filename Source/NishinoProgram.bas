Attribute VB_Name = "NishinoProgram"
'Public BOARD As String
'Public DEVICE As String
'Public RBUF As String
'Public OBUF As String
'Public Junk As String
'Public JUNKS As String
'Public ICH As String
'Public ACK As String
'Public CR As String
'Public DA As String
'Public DAd As Integer
'Public msg As String
'Public BD As Integer
'Public TEMP As String
'Public DV As Integer
'Public DU As Integer
'Public DW As Integer
'Public V As Double
'Public SAYOKO As Integer
'Public PBUF As String
'Public MonoStepsPerUnit As Double
'Public PBUFd As Integer
'Public WMONOPOS As Double
'Public MONOPOS As Double
'Public WaitTime As Variant
'Public Slit As Integer
'Public SlitPOS As Double
'Public WSlitPOS As Double
'Public SlitStepsPerUnit As Double
'Public Freq As Integer
'Public MONOSTEPSa As Double
'Public MONOSTEPSb As Double
'Public MONOSTEPSc As Double
'Public TE As Integer
'Public STARTW As Double
'Public STARTS As Double
'Public STOPW As Double
'Public STOPS As Double
'Public INTERVALW As Double
'Public INTERVALS As Double
'Public LF As String
'Public TEM As String
'Public COUNT As Integer
'Public INC As Integer
'Public SUGI As Integer
'
'
'
'
'
'
'Sub INITIALLIZE()
'
'MONOSTEPSa = -33.9449
'MONOSTEPSb = 0.0812
'MONOSTEPSc = -0.000000834
'MonoStepsPerUnit = 0.07
'SlitStepsPerUnit = 0.002
'BOARD = "GPIB0"
'DEVICE = "DEV"
'RBUF = Space(132)
'OBUF = Space(132)
'Junk = Space(132)
'ICH = Space(1)
'ACK = Space(1)
'CR = Chr(13)
'LF = Chr(10)
'
'DA = 1
'
'DEVICE = DEVICE & DA
'
'Call ibfind(BOARD, BD)  '�f�o�C�X�����s���ȂƂ��Ɏg�p����Ă��Ȃ��f�o�C�X���I�[�v�����ď��������s��'
'If BD < 0 Then
'TEMP = error()
'End If
'
'Call ibfind(DEVICE, DV) '������ 1'
'If DV < 0 Then
'TEMP = error()
'End If
'
'DA = 8
'DEVICE = "DEV"
'DEVICE = DEVICE & DA
'Call ibfind(DEVICE, DU) '���b�N�C���A���v 8'
'
'DA = 14
'DEVICE = "DEV"
'DEVICE = DEVICE & DA
'Call ibfind(DEVICE, DW) '�\�[�X���[�^�[ 14'
'
'EOSV = &HD
'V = EOSV + &H1400
'Call ibeos(DV, V)   '�I�[�����̐ݒ�'
'Call ibtmo(DV, T300ms)  '�^�C���A�E�g���Ԃ̐ݒ�'
'
'Call ibrd(DV, Junk)
'ICH = Chr(222)
'Call ibwrt(DV, ICH)
'
'Application.Wait (Now() + TimeValue("0:00:01"))
'
'Call ibrd(DV, Junk)
'
'ICH = " "
'Call ibwrt(DV, ICH)
'If (ibsta And EERR) Then
'TEMP = error()
'End If
'
'Call ibrd(DV, ICH)
'If (ibsta And EERR) Then
'TEMP = error()
'End If
'
'If ICH = "B" Then
'
'OBUF = "O2000" & Chr(0)
'Call ibwrt(DV, OBUF)
'If (ibsta% And EERR) Then
'TEMP = error()
'End If
'
'Application.Wait (Now() + TimeValue("0:00:01"))
'
'Call ibrd(DV, ICH)
'
'If (ibsta And EERR) Then
'TEMP = error()
'End If
'
'If ICH <> "*" Then
'Application.Wait (Now() + "0:00:01")
'End If
'
'End If
'
'ICH = " "
'Call ibwrt(DV, ICH)
'Call ibrd(DV, ICH)
'
'Cells(1, 1) = "Initiallized"
''MsgBox "IEEE-488 Communications Established!"
'SAYOK = 1   'ACK sound and OK print out flag'
'
'End Sub
'Sub MoveMotorRelative()
'
'WMONOPOS = InputBox("Move Motor Relative, How much ? nm�œ��́B", Title:="Move Moter Relative")
'MONOPOS = WMONOPOS / MonoStepsPerUnit   'MONOPOS�̓X�e�b�v���BWMONOPOS�͔g���\���B'
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'End Sub
'Sub MOVE_WORKING_ABS_POSITION()
'
'WMONOPOS = Cells(9, 4)
'
'If WMONOPOS > 1400 Or WMONOPOS < 0 Then
'MsgBox "�͈͊O"
'
'Else
'
'ICH = "Z61,1," & WMONOPOS & CR
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'
'ICH = "E"
'Call ibwrt(DV, ICH) '���[�^�[��busy�Ȃ�oq�Anot busy�Ȃ�oz��Ԃ�'
'Call ibrd(DV, RBUF)
'
'PBUF = Left(RBUF, InStr(RBUF, o) + 1) 'RBUF����o�̈ʒu����肵�Ao�̎��̕����܂ł�PBUF��'
'
'Cells(1, 1) = PBUF
'
'While PBUF = "oq" 'busy�Ȃ珈�����J��Ԃ�'
'ICH = "E" '���[�^�[��busy�Ȃ�oq�Anot busy�Ȃ�oz��Ԃ�'
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, o) + 1) 'RBUF����o�̈ʒu����肵�Ao�̎��̕����܂ł�PBUF��'
'Wend
'
'ICH = "Z62,1" & CR
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'
'Cells(8, 3) = "step"
'Cells(8, 4) = "wavelength(nm)"
'Cells(9, 2) = "present value"
'Cells(9, 4) = PBUF
'
'OBUF = "H0" & CR
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'MONOPOS = CInt(PBUF)
'
'Cells(9, 3) = PBUF
'
'End If
'
'End Sub
'Sub READ_WORKING_ABS_POSITION()
'
'ICH = "Z62,1" & CR
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'Cells(9, 4) = PBUF
'Cells(1, 1) = PBUF
'
'End Sub
'Sub Increase()
'
'ICH = "Z62,1" & CR
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'WMONOPOS = CInt(PBUF)
'
'INC = Cells(17, 4)
'
'ICH = "Z61,1," & WMONOPOS + INC & CR
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'
'Application.Wait (Now() + TimeValue("0:00:01"))
'Call READ_WORKING_ABS_POSITION
'
'End Sub
'
'Sub ReadMotorSpeed()
'
'OBUF = "C" & "0" & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'
'MsgBox "FREQMAX, FREQMIN, RAMPTIME" & Chr(13) & PBUF, 64, "Motor Speed"
'
'End Sub
'Function error() As String
'
'MsgBox ("�G���[�I")
'End Function
'Sub MotorInit()
'Call ibwrt(DV, "A")
'
'MsgBox "Wait 30 sec! Please Click [OK]."    '30�b�҂��Ȃ��ƃG���[���o��'
'Application.Wait (Now() + TimeValue("0:00:30"))
'MsgBox "OK"
'
'End Sub
'Sub FLASHVERSION()
'
'ICH = "y"
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'
'MsgBox "Boot Version: " & PBUF, 64, "Flash Version"
'
'ICH = "z"
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'
'MsgBox "Flash Version: " & PBUF, 64, "Flash Version"
'
'End Sub
'Sub MoterBusyCheck()
'
'ICH = "E" '���[�^�[��busy�Ȃ�oq�Anot busy�Ȃ�oz��Ԃ�'
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'Cells(1, 1) = RBUF
'
'End Sub
'
'Sub SetMonoPosition()
'
'WMONOPOS = InputBox("�g��(nm)����́B" & Chr(13) & "0(nm)�`1400(nm)�܂ŉ\�B", Title:="Set Mono Position")
'MONOPOS = (-MONOSTEPSb + Sqr(((MONOSTEPSb) ^ 2) - 4 * MONOSTEPSc * (MONOSTEPSa - WMONOPOS))) / (2 * MONOSTEPSc) '���̌����Ōv�Z'
'MONOPOS = Round(MONOPOS)    '�����Ɋۂ߂�'
'
'If WMONOPOS > 1400 Or WMONOPOS < 0 Then
'MsgBox "�͈͊O�ł�"
'
'Else
'OBUF = "H0" & CR
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'
'MONOPOS = MONOPOS - CInt(PBUF)
'
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'End If
'
'End Sub
'Sub SetMonoPosition2()
'
'WMONOPOS = Cells(9, 4)
'MONOPOS = (-MONOSTEPSb + Sqr(((MONOSTEPSb) ^ 2) - 4 * MONOSTEPSc * (MONOSTEPSa - WMONOPOS))) / (2 * MONOSTEPSc) '���̌����Ōv�Z'
'MONOPOS = Round(MONOPOS)    '�����Ɋۂ߂�'
'
'If WMONOPOS > 1400 Or WMONOPOS < 0 Then
'MsgBox "�͈͊O�ł�"
'
'Else
'OBUF = "H0" & CR '���݂̃X�e�b�v����ǂ�'
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1) 'RBUF����CR�̈ʒu����肵�ACR��荶�̕�����PBUF'
'PBUF = Right(PBUF, Len(PBUF) - 1)
'
'MONOPOS = MONOPOS - CInt(PBUF)
'
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13) '���[�^�[�𓮂���'
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'ICH = "E"
'Call ibwrt(DV, ICH) '���[�^�[��busy�Ȃ�oq�Anot busy�Ȃ�oz��Ԃ�'
'Call ibrd(DV, RBUF)
'
'PBUF = Left(RBUF, InStr(RBUF, o) + 1) 'RBUF����o�̈ʒu����肵�Ao�̎��̕����܂ł�PBUF��'
'
'Cells(1, 1) = PBUF
'
'While PBUF = "oq" 'busy�Ȃ珈�����J��Ԃ�'
'ICH = "E" '���[�^�[��busy�Ȃ�oq�Anot busy�Ȃ�oz��Ԃ�'
'Call ibwrt(DV, ICH)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, o) + 1) 'RBUF����o�̈ʒu����肵�Ao�̎��̕����܂ł�PBUF��'
'Wend
'
'OBUF = "H0" & CR '���݂̃X�e�b�v����ǂ�'
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'MONOPOS = CInt(PBUF)
'
'Cells(9, 2) = "present value"
'Cells(11, 2) = "start value"
'Cells(12, 2) = "stop value"
'Cells(9, 3) = MONOPOS
'Cells(9, 4) = Round(MONOSTEPSa + (MONOPOS * MONOSTEPSb) + (MONOPOS * MONOPOS * MONOSTEPSc))
'
'End If
'
'End Sub
'
'Sub READMONOPOSITION()
'
'OBUF = "H0" & CR
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'MONOPOS = CInt(PBUF)
'
'Cells(9, 3) = MONOPOS
'Cells(9, 4) = Round(MONOSTEPSa + (MONOPOS * MONOSTEPSb) + (MONOPOS * MONOPOS * MONOSTEPSc))
'MsgBox "Steps: " & MONOPOS & Chr(13) & "wave length: " & Round(MONOSTEPSa + (MONOPOS * MONOSTEPSb) + (MONOPOS * MONOPOS * MONOSTEPSc)) & " (nm)", 64, "Mono Position"
'
'End Sub
'Sub HARDWARESTATUS() '�R���g���[���Ȃ�Ƃ������̂������ĂȂ�����g���Ȃ�'
'
'ICH = "r"
'Call ibwrt(DV, ICH)
'
'Call ibrd(DV, RBUF)
'
'MsgBox RBUF, 64, "Hard Ware Status"
'
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'
'MsgBox PBUF, 64, "Hard Ware Status"
'
'End Sub
'Sub GoToRS232()
'
'MsgBox "Change to Low IQ Mode"
'Call ibwrt(DV, "Y")
'Call ibloc(DV)
'MsgBox "Change Switch to Hand Held Position." & Chr(13) & "then Hit <.> on Hand Held twice"
'
'End Sub
'Sub ComeBackToIEE488()
'
'MsgBox "IEEE-488 coming On-line" & Chr(13) & "RS-232 I/O should be SILENT"
'
'End Sub
'Sub SlitSetSpeed()
'
'Slit = InputBox("Slit ?", Title:="Slit Set Speed")
'Freq = InputBox("Freqency ?", Title:="Slit Set Speed")
'OBUF = "g0," & Str(Slit) & "," & Str(Freq) & Chr(13)
'Call ibwrt(DV, OBUF)
'
'End Sub
'Sub SlitReadSpeed()
'
'Slit = InputBox("Slit ?", Title:="Slit Set Speed")
'OBUF = "h0," & Str(Slit) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
'
'MsgBox "Slit Speed: " & PBUF, 64, "Slit Read Speed"
'
'End Sub
'Sub SlitSetPosition() '�X���b�g�ԍ����ʂɎw�肵�ăX���b�g�𓮂���'
'
'Slit = InputBox("�X���b�g�ԍ��̓���" & Chr(13) & "0, 2, 3�̐����œ��́B", Title:="Slit Set Position") '�X���b�g1�͖���'
'
'SlitPSP = -1
'Do While (SlitPOS > 1120 Or SlitPOS < 0)
'WSlitPOS = InputBox("�X���b�g��(mm)�̓��́B" & Chr(13) _
'& "0(mm)����2(mm)�܂ŉ\�B", Title:="Set Slit Position")
'SlitPOS = WSlitPOS / SlitStepsPerUnit
'Loop
'
'OBUF = "j0," & Str(Slit) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
'
'SlitPOS = SlitPOS - CInt(PBUF)
'
'OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'End Sub
'
'Sub SlitSetPosition2() '�Z������l��ǂݎ���ăX���b�g3�𓮂���'
'
'Slit = 0
'While Slit < 4
'
'WSlitPOS = Cells(9 + Slit, 9)
'If WSlitPOS > 2 Then
'
'MsgBox "�G���[�I"
'
'Else
'SlitPOS = WSlitPOS / SlitStepsPerUnit
'
'OBUF = "j0," & Str(Slit) & Chr(13) '���݂̃X���b�g�ʒu��ǂ�'
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
'
'SlitPOS = SlitPOS - CInt(PBUF) '���ΓI�ɂ����瓮�����΂������v�Z'
'
'OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13) '�X�e�b�v�œ�����'
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'End If
'Slit = Slit + 1
'Application.Wait (Now() + TimeValue("0:00:02"))
'Wend
'
'Call SlitReadPosition
'
'
'End Sub
'Sub slittest()
'
'TE = Cells(1, 1)
'
'Cells(1, 2) = TE
'
'
'End Sub
'
'Sub SlitSet3()  '�X���b�g3�𓯎��ɓ������������ǂ܂������Ȃ�'
'
'Do While (SlitPOS > 1120 Or SlitPOS < 0)
'WSlitPOS = InputBox("�X���b�g����(mm)�œ��́B" & Chr(13) _
'& "0(mm)����2(mm)�܂ŉ\�B", Title:="Set Slit Position")
'SlitPOS = WSlitPOS / SlitStepsPerUnit
'Loop
'
'Slit = 0
'
'Do While (Slit > 3) '���������^�̊Ԃ͏������J��Ԃ�'
'OBUF = "j0," & Str(Slit) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
'
'SlitPOS = SlitPOS - CInt(PBUF)
'
'OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'Slit = Slit + 1
'Loop
'
'End Sub
'Sub SlitReadPosition()
'
'Slit = 0
'
'While Slit < 4
'
'OBUF = "j0," & Str(Slit) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
'SlitPOS = CInt(PBUF)
'
'Cells(8, 8) = "Step"
'Cells(8, 9) = "Width(nm)"
'Cells(9 + Slit, 7) = "Slit" & Slit
'Cells(9 + Slit, 8) = SlitPOS
'Cells(9 + Slit, 9) = SlitPOS * SlitStepsPerUnit
'
'Slit = Slit + 1
'Wend
'
'SlitPOS = -1
'
'End Sub
'Sub SlitMoveRelative()
'
'Slit = InputBox("�X���b�g�ԍ��̓���" & Chr(13) & "0����3�܂ł̐����œ��́B", Title:="Slit Move Relative")
'SlitPOS = InputBox("How much ? �X�e�b�v�������", Title:="Slit Move Relative")
'
'
'OBUF = "k0," & Str(Slit) & "," & Str(SlitPOS) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'End Sub
'Sub test()
'
'Dim WMONOPOS As Double
'Dim MONOPOS As Double
'
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)   'PBUF�̈�ԍ��̕���������'
'MONOPOS = CInt(PBUF)    '�����^�ɕϊ�'
'
'Do While MONOPOS < 0
'
'WMONOPOS = 3
'MONOPOS = WMONOPOS / MonoStepsPerUnit
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'OBUF = "H0" & CR
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'MONOPOS = CInt(PBUF)
'MsgBox "Steps: " & MONOPOS & Chr(13) & Chr(10) & "wave length: " & MONOPOS * MonoStepsPerUnit, 64, "Mono Position"
'
'Loop
'
'End Sub
'Sub step()
'
'Do While (MONOPOS > 20000 Or MONOPOS < -1457)
'MONOPOS = InputBox("�ǂ��Ɉړ����܂����H�X�e�b�v���œ��́B" & Chr(13) _
'& "-1457(step)�`20000(step)�܂ŉ\�B", Title:="Set Mono Position")
'Loop
'
'OBUF = "H0" & CR
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'
'MONOPOS = MONOPOS - CInt(PBUF)
'
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13)
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'MONOPOS = -1458
'
'End Sub
'Sub test2()
'
'Call ibwrt(DV, " ")
'
'Call ibrd(DV, ICH)
'MsgBox ICH
'
'Call ibrd(DV, ACK)
'
'Cells(1, 1) = ACK
'
'If ACK = "o" Then
'MsgBox "Receive was O.K."
'End If
'
'If ACK = "b" Then
'MsgBox "Receive was BAD"
'End If
'
'End Sub
'
'Sub lockin()
'
'DEVICE = "DEV"
'DA = 3
'DAd = CInt(DA)  '�����^(Integer)�ւ̃f�[�^�ϊ�'
'CR = Chr(13)
'
'DEVICE = DEVICE & DA
'
'Call ibfind(BOARD, BD)
'Call ibfind(DEVICE, DU)
'
'ICH = "OUTR? 1"
'
'Call ibwrt(DU, ICH)
'
'Application.Wait (Now() + TimeValue("0:00:00"))
'
'Call ibrd(DU, Junk)
'
'Cells(25, 11) = "channel1"
'Cells(26, 11) = Junk
'Cells(26, 12) = "mV"
'
'End Sub
'
'Sub sokutei()
'
'I = 1
'
'While I < 20
'
'MONOPOS = 20 '20�X�e�b�v������'
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13)
'Call ibwrt(DV, OBUF) '������𓮂�������'
'Call ibrd(DV, RBUF) '���[�h���Ȃ��Ɠ����Ȃ�'
'
'Application.Wait (Now() + TimeValue("0:00:01"))
'
'Call ibwrt(DU, "OUTR? 1") '���b�N�C���A���v�̒l��ǂޖ���'
'Call ibrd(DU, Junk) 'JUNK�ɓ����'
'
'OBUF = "H0" & CR
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'MONOPOS = CInt(PBUF)
'
'Cells(1, 11) = "wavelength(nm)"
'Cells(1, 12) = "intensity"
'Cells(1 + I, 11) = Round(MONOSTEPSa + (MONOPOS * MONOSTEPSb) + (MONOPOS * MONOPOS * MONOSTEPSc))
'Cells(1 + I, 12) = Junk
'I = I + 1
'
'Wend
'
'End Sub
'
'Sub sokutei2()
'
'STARTW = Cells(11, 4)
'STOPW = Cells(12, 4)
'INTERVALW = Cells(13, 4)
'STARTS = (-MONOSTEPSb + Sqr(((MONOSTEPSb) ^ 2) - 4 * MONOSTEPSc * (MONOSTEPSa - STARTW))) / (2 * MONOSTEPSc) '���̌����Ōv�Z'
'STARTS = Round(STARTS)    '�����Ɋۂ߂�'
'STOPS = (-MONOSTEPSb + Sqr(((MONOSTEPSb) ^ 2) - 4 * MONOSTEPSc * (MONOSTEPSa - STOPW))) / (2 * MONOSTEPSc) '���̌����Ōv�Z'
'STOPS = Round(STOPS)    '�����Ɋۂ߂�'
'
'
'WMONOPOS = STARTW
'
'While WMONOPOS < STOPW + 1
'
'MONOPOS = (-MONOSTEPSb + Sqr(((MONOSTEPSb) ^ 2) - 4 * MONOSTEPSc * (MONOSTEPSa - WMONOPOS))) / (2 * MONOSTEPSc) '�ڕW�g���̃X�e�b�v�����v�Z'
'MONOPOS = Round(MONOPOS)    '�����Ɋۂ߂�'
'
'OBUF = "H0" & CR '���݂̃X�e�b�v����ǂ�'
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1) 'RBUF����CR�̈ʒu����肵�ACR��荶�̕�����PBUF'
'PBUF = Right(PBUF, Len(PBUF) - 1) '�����̒�������肵�A���[��1�����ȊO��PBUF�ɂ���'
'
'MONOPOS = MONOPOS - CInt(PBUF) '�ڕW�g���܂ł̃X�e�b�v����MONOPOS�ɓ����'
'
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13) '���[�^�[�𓮂���'
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'
'OBUF = "F" & "0" & "," & Str(MONOPOS) & Chr(13)
'Call ibwrt(DV, OBUF) '������𓮂�������'
'Call ibrd(DV, RBUF) '���[�h���Ȃ��Ɠ����Ȃ�'
'
'Application.Wait (Now() + TimeValue("0:00:01"))
'
'Call ibwrt(DU, "OUTR? 1") '���b�N�C���A���v�̒l��ǂޖ���'
'Call ibrd(DU, Junk) 'JUNK�ɓ����'
'
'OBUF = "H0" & CR
'Call ibwrt(DV, OBUF)
'Call ibrd(DV, RBUF)
'PBUF = Left(RBUF, InStr(RBUF, CR) - 1)
'PBUF = Right(PBUF, Len(PBUF) - 1)
'MONOPOS = CInt(PBUF)
'
'Cells(1, 11) = "wavelength(nm)"
'Cells(1, 12) = "intensity"
'Cells(2 + I, 11) = Round(MONOSTEPSa + (MONOPOS * MONOSTEPSb) + (MONOPOS * MONOPOS * MONOSTEPSc))
'Cells(2 + I, 12) = Junk
'I = I + 1
'
'WMONOPOS = WMONOPOS + INTERVALW
'
'Wend
'
'End Sub
'Sub MEASURE()
'
'I = 0
'
'STARTS = Cells(17, 8)
'STOPS = Cells(18, 8)
'INTERVALS = Cells(19, 8)
'
'While I < 110
'
'ICH = "*RST" 'GPIB�f�t�H���g����'
'Call ibwrt(DW, ICH)
'ICH = ":SENS:FUNC:CONC OFF" '�����@�\���I�t'
'Call ibwrt(DW, ICH)
'
'ICH = ":SOUR:FUNC VOLT" '�d���\�[�X��I������'
'Call ibwrt(DW, ICH)
'
'ICH = ":SOUR:VOLT:MODE FIXED" '�Œ�d���\�[�X���[�h'
'Call ibwrt(DW, ICH)
'
'ICH = ":SOUR:VOLT:RANG 200" '20V�\�[�X�����W��I��'
'Call ibwrt(DW, ICH)
'
'ICH = ":SOUR:VOLT:LEV " & 2 * I '�\�[�X�o�́�10V'
'Call ibwrt(DW, ICH)
'
'ICH = ":SENS:CURR:PROT 10E-3" '�R���v���C�A���X'
'Call ibwrt(DW, ICH)
'
'
'ICH = ":OUTP ON" '�o�͂�ON'
'Call ibwrt(DW, ICH)
'
'Application.Wait (Now() + TimeValue("0:00:10"))
'
'Call ibwrt(DU, "OUTR? 1") '���b�N�C���A���v�̒l��ǂޖ���'
'Call ibrd(DU, Junk) 'JUNK�ɓ����'
'Junk = CVar(Junk)
'
'Cells(1, 12) = "intensity(V)"
'Cells(I + 16, 12) = Junk
'
'I = I + 10
'Wend
'
'End Sub
'
'Sub SorcematerTEST()
'
'STARTS = Cells(17, 8)
'STOPS = Cells(18, 8)
'INTERVALS = Cells(19, 8)
'
'
'ICH = "*RST" 'GPIB�f�t�H���g����'
'Call ibwrt(DW, ICH)
'ICH = ":SENS:FUNC:CONC OFF" '�����@�\���I�t'
'Call ibwrt(DW, ICH)
'
'ICH = ":SOUR:FUNC VOLT" '�\�[�XV'
'Call ibwrt(DW, ICH)
'ICH = ":SENS:FUNC 'CURR'" '���W���[I'
'Call ibwrt(DW, ICH)
'
''ICH = ":SENS:FUNC VOLT"
''Call ibwrt(DW, ICH)
''ICH = ":SENS:FUNC 'CURR:DC'"
''Call ibwrt(DW, ICH)
'ICH = ":SENS:CURR:PROT 0.1" '�d���R���v���C�A���X100mA'
'Call ibwrt(DW, ICH)
'ICH = ":SOUR:VOLT:START " & STARTS  '�J�n�d��'
'Call ibwrt(DW, ICH)
'ICH = ":SOUR:VOLT:STOP " & STOPS  '��~�d��'
'Call ibwrt(DW, ICH)
'ICH = ":SOUR:VOLT:STEP " & INTERVALS  '�X�e�b�v�d��'
'Call ibwrt(DW, ICH)
'ICH = ":SOUR:VOLT:MODE SWE" '�d���X�C�[�v���[�h��I��'
'Call ibwrt(DW, ICH)
'COUNT = (STOPS - STARTS) / INTERVALS + 1
'ICH = ":TRIG:COUN " & COUNT '�g���K�J�E���g=�_���@�_��=(��~-�J�n)/�X�e�b�v��+1'
'Call ibwrt(DW, ICH)
'ICH = ":SOUR:DEL 0.05" '�\�[�X�f�B���C'
'Call ibwrt(DW, ICH)
'ICH = ":FORMat:ELEMents VOLTage, CURRent " '�o�b�t�@�ɏ������ނ̂�V��I'
'Call ibwrt(DW, ICH)
'ICH = ":OUTPUT ON" '�\�[�X�o�͂��I����'
'Call ibwrt(DW, ICH)
'ICH = ":READ?" '�X�C�[�v���g���K���A�f�[�^�𐿋����Ă�������'
'Call ibwrt(DW, ICH)
'
'TEM = COUNT * 28 'V��I�̃����Z�b�g��28����
'TEM = CInt(TEM)
'TEM = Space(TEM) '�󂯎���X�y�[�X���m��
'Call ibrd(DW, TEM)
'
'Cells(1, 1) = TEM
'
'I = 1
'Cells(1, 13) = "Voltage(V)"
'Cells(1, 14) = "Current(A)"
'
'While I < COUNT * 2 + 1
'
''13�����łЂƂ̐��l�f�[�^�B�R���}������Ŏ��̐��l�f�[�^�������B
'
'Junk = Left(TEM, 13) 'TEM�̍�13������JUNK�ɓ����'
'TEM = Right(TEM, Len(TEM) - 14) '��14�����������B'
'
'If I Mod 2 = 1 Then
'Cells(I - ((I - 1) / 2) + 1, 13) = Junk
'Else
'Cells(I / 2 + 1, 14) = Junk
'End If
'I = I + 1
'Wend
'
'ICH = ":OUTP OFF" '�o�̓I�t'
'Call ibwrt(DW, ICH)
'
'End Sub
'
'Sub Buffer_Clear()
'
'TEM = COUNT * 28
'TEM = CInt(TEM)
'TEM = Space(TEM)
'Call ibrd(DW, TEM)
'
'I = 1
'
'While I < COUNT * 2 + 1
'
'Junk = Left(TEM, 13) 'TEM�̍�13������JUNK�ɓ����'
'TEM = Right(TEM, Len(TEM) - 14) '��14�����������B'
'
'If I Mod 2 = 1 Then
'Cells(I - ((I - 1) / 2), 13) = Junk
'
'Else
'Cells(I / 2, 14) = Junk
'
'End If
'
'I = I + 1
'Wend
'
'End Sub
'
'
