Attribute VB_Name = "VBIB32"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 32-bit Visual Basic Language Interface
' Version 1.81
' Copyright 2001 National Instruments Corporation.
' All Rights Reserved.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This module contains the subroutine declarations,
'   function declarations and constants required to use
'   the National Instruments GPIB Dynamic Link Library
'   (DLL) for controlling IEEE-488 instrumentation.  This
'   file must be 'added' to your Visual Basic project
'   (by choosing Add File from the File menu or pressing
'   CTRL+F12) so that you can access the NI-488.2
'   subroutines and functions.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NI-488.2 DLL entry function declarations

Declare Function ibask32 Lib "Gpib-32.dll" Alias "ibask" (ByVal UD As Long, ByVal opt As Long, value As Long) As Long
Declare Function ibbna32 Lib "Gpib-32.dll" Alias "ibbnaA" (ByVal UD As Long, sstr As Any) As Long
Declare Function ibcac32 Lib "Gpib-32.dll" Alias "ibcac" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibclr32 Lib "Gpib-32.dll" Alias "ibclr" (ByVal UD As Long) As Long
Declare Function ibcmd32 Lib "Gpib-32.dll" Alias "ibcmd" (ByVal UD As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibcmda32 Lib "Gpib-32.dll" Alias "ibcmda" (ByVal UD As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibconfig32 Lib "Gpib-32.dll" Alias "ibconfig" (ByVal UD As Long, ByVal opt As Long, ByVal V As Long) As Long
Declare Function ibdev32 Lib "Gpib-32.dll" Alias "ibdev" (ByVal bdid As Long, ByVal pad As Long, ByVal sad As Long, ByVal tmo As Long, ByVal eot As Long, ByVal eos As Long) As Long
Declare Function ibdma32 Lib "Gpib-32.dll" Alias "ibdma" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibeos32 Lib "Gpib-32.dll" Alias "ibeos" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibeot32 Lib "Gpib-32.dll" Alias "ibeot" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibfind32 Lib "Gpib-32.dll" Alias "ibfindA" (sstr As Any) As Long
Declare Function ibgts32 Lib "Gpib-32.dll" Alias "ibgts" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibist32 Lib "Gpib-32.dll" Alias "ibist" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function iblck32 Lib "Gpib-32.dll" Alias "iblck" (ByVal UD As Long, ByVal V As Long, ByVal LockWaitTime As Long, arg1 As Any) As Long
Declare Function iblines32 Lib "Gpib-32.dll" Alias "iblines" (ByVal UD As Long, V As Long) As Long
Declare Function ibln32 Lib "Gpib-32.dll" Alias "ibln" (ByVal UD As Long, ByVal pad As Long, ByVal sad As Long, ln As Long) As Long
Declare Function ibloc32 Lib "Gpib-32.dll" Alias "ibloc" (ByVal UD As Long) As Long
Declare Function iblock32 Lib "Gpib-32.dll" Alias "iblock" (ByVal UD As Long) As Long
Declare Function ibonl32 Lib "Gpib-32.dll" Alias "ibonl" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibpad32 Lib "Gpib-32.dll" Alias "ibpad" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibpct32 Lib "Gpib-32.dll" Alias "ibpct" (ByVal UD As Long) As Long
Declare Function ibppc32 Lib "Gpib-32.dll" Alias "ibppc" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibrd32 Lib "Gpib-32.dll" Alias "ibrd" (ByVal UD As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrda32 Lib "Gpib-32.dll" Alias "ibrda" (ByVal UD As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrdf32 Lib "Gpib-32.dll" Alias "ibrdfA" (ByVal UD As Long, sstr As Any) As Long
Declare Function ibrpp32 Lib "Gpib-32.dll" Alias "ibrpp" (ByVal UD As Long, sstr As Any) As Long
Declare Function ibrsc32 Lib "Gpib-32.dll" Alias "ibrsc" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibrsp32 Lib "Gpib-32.dll" Alias "ibrsp" (ByVal UD As Long, sstr As Any) As Long
Declare Function ibrsv32 Lib "Gpib-32.dll" Alias "ibrsv" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibsad32 Lib "Gpib-32.dll" Alias "ibsad" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibsic32 Lib "Gpib-32.dll" Alias "ibsic" (ByVal UD As Long) As Long
Declare Function ibsre32 Lib "Gpib-32.dll" Alias "ibsre" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibstop32 Lib "Gpib-32.dll" Alias "ibstop" (ByVal UD As Long) As Long
Declare Function ibtmo32 Lib "Gpib-32.dll" Alias "ibtmo" (ByVal UD As Long, ByVal V As Long) As Long
Declare Function ibtrg32 Lib "Gpib-32.dll" Alias "ibtrg" (ByVal UD As Long) As Long
Declare Function ibunlock32 Lib "Gpib-32.dll" Alias "ibunlock" (ByVal UD As Long) As Long
Declare Function ibwait32 Lib "Gpib-32.dll" Alias "ibwait" (ByVal UD As Long, ByVal mask As Long) As Long
Declare Function ibwrt32 Lib "Gpib-32.dll" Alias "ibwrt" (ByVal UD As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrta32 Lib "Gpib-32.dll" Alias "ibwrta" (ByVal UD As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrtf32 Lib "Gpib-32.dll" Alias "ibwrtfA" (ByVal UD As Long, sstr As Any) As Long
Declare Sub AllSpoll32 Lib "Gpib-32.dll" Alias "AllSpoll" (ByVal boardID As Long, arg1 As Any, arg2 As Any)
Declare Sub DevClear32 Lib "Gpib-32.dll" Alias "DevClear" (ByVal boardID As Long, ByVal V As Long)
Declare Sub DevClearList32 Lib "Gpib-32.dll" Alias "DevClearList" (ByVal boardID As Long, arg1 As Any)
Declare Sub EnableLocal32 Lib "Gpib-32.dll" Alias "EnableLocal" (ByVal boardID As Long, arg1 As Any)
Declare Sub EnableRemote32 Lib "Gpib-32.dll" Alias "EnableRemote" (ByVal boardID As Long, arg1 As Any)
Declare Sub FindLstn32 Lib "Gpib-32.dll" Alias "FindLstn" (ByVal boardID As Long, arg1 As Any, arg2 As Any, ByVal limit As Long)
Declare Sub FindRQS32 Lib "Gpib-32.dll" Alias "FindRQS" (ByVal boardID As Long, arg1 As Any, result As Long)
Declare Sub PassControl32 Lib "Gpib-32.dll" Alias "PassControl" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub PPoll32 Lib "Gpib-32.dll" Alias "PPoll" (ByVal boardID As Long, result As Long)
Declare Sub PPollConfig32 Lib "Gpib-32.dll" Alias "PPollConfig" (ByVal boardID As Long, ByVal addr As Long, ByVal line As Long, ByVal sense As Long)
Declare Sub PPollUnconfig32 Lib "Gpib-32.dll" Alias "PPollUnconfig" (ByVal boardID As Long, arg1 As Any)
Declare Sub RcvRespMsg32 Lib "Gpib-32.dll" Alias "RcvRespMsg" (ByVal boardID As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReadStatusByte32 Lib "Gpib-32.dll" Alias "ReadStatusByte" (ByVal boardID As Long, ByVal addr As Long, result As Long)
Declare Sub Receive32 Lib "Gpib-32.dll" Alias "Receive" (ByVal boardID As Long, ByVal addr As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReceiveSetup32 Lib "Gpib-32.dll" Alias "ReceiveSetup" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub ResetSys32 Lib "Gpib-32.dll" Alias "ResetSys" (ByVal boardID As Long, arg1 As Any)
Declare Sub Send32 Lib "Gpib-32.dll" Alias "Send" (ByVal boardID As Long, ByVal addr As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendCmds32 Lib "Gpib-32.dll" Alias "SendCmds" (ByVal boardID As Long, sstr As Any, ByVal cnt As Long)
Declare Sub SendDataBytes32 Lib "Gpib-32.dll" Alias "SendDataBytes" (ByVal boardID As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendIFC32 Lib "Gpib-32.dll" Alias "SendIFC" (ByVal boardID As Long)
Declare Sub SendList32 Lib "Gpib-32.dll" Alias "SendList" (ByVal boardID As Long, arg1 As Any, arg2 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendLLO32 Lib "Gpib-32.dll" Alias "SendLLO" (ByVal boardID As Long)
Declare Sub SendSetup32 Lib "Gpib-32.dll" Alias "SendSetup" (ByVal boardID As Long, arg1 As Any)
Declare Sub SetRWLS32 Lib "Gpib-32.dll" Alias "SetRWLS" (ByVal boardID As Long, arg1 As Any)
Declare Sub TestSys32 Lib "Gpib-32.dll" Alias "TestSys" (ByVal boardID As Long, arg1 As Any, arg2 As Any)
Declare Sub Trigger32 Lib "Gpib-32.dll" Alias "Trigger" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub TriggerList32 Lib "Gpib-32.dll" Alias "TriggerList" (ByVal boardID As Long, arg1 As Any)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIB global variables

Declare Function RegisterGpibGlobalsForThread Lib "Gpib-32.dll" (Longibsta As Long, Longiberr As Long, Longibcnt As Long, ibcntl As Long) As Long
Declare Function UnregisterGpibGlobalsForThread Lib "Gpib-32.dll" () As Long
Declare Function ThreadIbsta32 Lib "Gpib-32.dll" Alias "ThreadIbsta" () As Long
Declare Function ThreadIbcnt32 Lib "Gpib-32.dll" Alias "ThreadIbcnt" () As Long
Declare Function ThreadIbcntl32 Lib "Gpib-32.dll" Alias "ThreadIbcntl" () As Long
Declare Function ThreadIberr32 Lib "Gpib-32.dll" Alias "ThreadIberr" () As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIBnotify OLE control

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIB-ENET functions

Declare Function iblockx32 Lib "Gpib-32.dll" Alias "iblockxA" (ByVal UD As Long, ByVal LockWaitTime As Long, arg1 As Any) As Long
Declare Function ibunlockx32 Lib "Gpib-32.dll" Alias "ibunlockx" (ByVal UD As Long) As Long


Sub AllSpoll(ByVal boardID As Integer, addrs() As Integer, results() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call AllSpoll32(boardID, addrs(0), results(0))

    Call copy_ibvars
End Sub

Sub copy_ibvars()
    ibsta = ConvertLongToInt(Longibsta)
    iberr = CInt(Longiberr)
    ibcnt = ConvertLongToInt(ibcntl)
End Sub

Sub DevClear(ByVal boardID As Integer, ByVal addr As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call DevClear32(boardID, addr)

    Call copy_ibvars
End Sub

Sub DevClearList(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call DevClearList32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub EnableLocal(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call EnableLocal32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub EnableRemote(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call EnableRemote32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub FindLstn(ByVal boardID As Integer, addrs() As Integer, results() As Integer, ByVal limit As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call FindLstn32(boardID, addrs(0), results(0), limit)

    Call copy_ibvars
End Sub

Sub FindRQS(ByVal boardID As Integer, addrs() As Integer, result As Integer)
   Dim tmpresult As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call FindRQS32(boardID, addrs(0), tmpresult)

    result = ConvertLongToInt(tmpresult)

    Call copy_ibvars
End Sub

Sub ibask(ByVal UD As Integer, ByVal opt As Integer, rval As Integer)
  Dim tmprval As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibask32(UD, opt, tmprval)

    rval = ConvertLongToInt(tmprval)

    Call copy_ibvars
End Sub

Sub ibbna(ByVal UD As Integer, ByVal udname As String)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibbna32(UD, ByVal udname)

    Call copy_ibvars
End Sub

Sub ibcac(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibcac32(UD, V)

    Call copy_ibvars
End Sub

Sub ibclr(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibclr32(UD)

    Call copy_ibvars
End Sub

Sub ibcmd(ByVal UD As Integer, ByVal buf As String)
   Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call ibcmd32(UD, ByVal buf, cnt)

    Call copy_ibvars
End Sub

Sub ibcmda(ByVal UD As Integer, ByVal buf As String)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call ibcmd32(UD, ByVal buf, cnt)

' When Visual Basic remapping buffer problem solved, then use:
'    call ibcmda32(ud, ByVal buf, cnt)

    Call copy_ibvars
End Sub

Sub ibconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibconfig32(bdid, opt, V)

    Call copy_ibvars
End Sub

Sub ibdev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer, UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    UD = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

    Call copy_ibvars
End Sub

Sub ibdma(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibdma32(UD, V)

    Call copy_ibvars
End Sub

Sub ibeos(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibeos32(UD, V)

    Call copy_ibvars
End Sub

Sub ibeot(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibeot32(UD, V)

    Call copy_ibvars
End Sub

Sub ibfind(ByVal udname As String, UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    UD = ConvertLongToInt(ibfind32(ByVal udname))

    Call copy_ibvars
End Sub

Sub ibgts(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibgts32(UD, V)

    Call copy_ibvars
End Sub

Sub ibist(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibist32(UD, V)

    Call copy_ibvars
End Sub

Sub iblines(ByVal UD As Integer, lines As Integer)
   Dim tmplines As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call iblines32(UD, tmplines)

    lines = ConvertLongToInt(tmplines)

    Call copy_ibvars
End Sub

Sub ibln(ByVal UD As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer)
    Dim tmpln As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibln32(UD, pad, sad, tmpln)

    ln = ConvertLongToInt(tmpln)

    Call copy_ibvars
End Sub

Sub ibloc(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibloc32(UD)

    Call copy_ibvars
End Sub

Sub iblck(ByVal UD As Integer, ByVal V As Integer, ByVal LockWaitTime As Long)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call iblck32(UD, V, LockWaitTime, ByVal 0)

    Call copy_ibvars
End Sub

Sub ibonl(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibonl32(UD, V)

    Call copy_ibvars
End Sub

Sub ibpad(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibpad32(UD, V)

    Call copy_ibvars
End Sub

Sub ibpct(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibpct32(UD)

    Call copy_ibvars
End Sub

Sub ibppc(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibppc32(UD, V)

    Call copy_ibvars
End Sub

Sub ibrd(ByVal UD As Integer, buf As String)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call ibrd32(UD, ByVal buf, cnt)

    Call copy_ibvars
End Sub

Sub ibrda(ByVal UD As Integer, buf As String)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call ibrd32(UD, ByVal buf, cnt)

' When Visual Basic remapping buffer problem solved, use this:
'    Call ibrda32(ud, ByVal buf, cnt)

    Call copy_ibvars
End Sub

Sub ibrdf(ByVal UD As Integer, ByVal filename As String)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibrdf32(UD, ByVal filename)

    Call copy_ibvars
End Sub

Sub ibrdi(ByVal UD As Integer, ibuf() As Integer, ByVal cnt As Long)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibrd32(UD, ibuf(0), cnt)

    Call copy_ibvars
End Sub

Sub ibrdia(ByVal UD As Integer, ibuf() As Integer, ByVal cnt As Long)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibrd32(UD, ibuf(0), cnt)

' When Visual Basic remapping buffer problem is solved, then use:
'    Call ibrda32(u, ibuf(0), cnt)

    Call copy_ibvars
End Sub

Sub ibrpp(ByVal UD As Integer, ppr As Integer)
    Static tmp_str As String * 2

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibrpp32(UD, ByVal tmp_str)

    ppr = Asc(tmp_str)

    Call copy_ibvars
End Sub

Sub ibrsc(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibrsc32(UD, V)

    Call copy_ibvars
End Sub

Sub ibrsp(ByVal UD As Integer, spr As Integer)
    Static tmp_str As String * 2

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL
    Call ibrsp32(UD, ByVal tmp_str)

    spr = Asc(tmp_str)

    Call copy_ibvars
End Sub

Sub ibrsv(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibrsv32(UD, V)

    Call copy_ibvars
End Sub

Sub ibsad(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibsad32(UD, V)

    Call copy_ibvars
End Sub

Sub ibsic(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibsic32(UD)

    Call copy_ibvars
End Sub

Sub ibsre(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibsre32(UD, V)

    Call copy_ibvars
End Sub

Sub ibstop(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibstop32(UD)

    Call copy_ibvars
End Sub

Sub ibtmo(ByVal UD As Integer, ByVal V As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibtmo32(UD, V)

    Call copy_ibvars
End Sub

Sub ibtrg(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call 32-bit DLL.
    Call ibtrg32(UD)

    Call copy_ibvars
End Sub

Sub ibwait(ByVal UD As Integer, ByVal mask As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibwait32(UD, mask)

    Call copy_ibvars
End Sub

Sub ibwrt(ByVal UD As Integer, ByVal buf As String)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call ibwrt32(UD, ByVal buf, cnt)

    Call copy_ibvars
End Sub

Sub ibwrta(ByVal UD As Integer, ByVal buf As String)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call ibwrt32(UD, ByVal buf, cnt)

' When Visual Basic remapping buffer problem is solved, use this:
'    Call ibwrta32(ud, ByVal buf, cnt)

    Call copy_ibvars
End Sub

Sub ibwrtf(ByVal UD As Integer, ByVal filename As String)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibwrtf32(UD, ByVal filename)

    Call copy_ibvars
End Sub

Sub ibwrti(ByVal UD As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibwrt32(UD, ibuf(0), cnt)

    Call copy_ibvars
End Sub

Sub ibwrtia(ByVal UD As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibwrt32(UD, ibuf(0), cnt)

' When Visual Basic remapping buffer problem is solved, use this:
'    Call ibwrta32(ud, ibuf(0), cnt)

    Call copy_ibvars
End Sub

Function ilask(ByVal UD As Integer, ByVal opt As Integer, rval As Integer) As Integer
    Dim tmprval As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilask = ConvertLongToInt(ibask32(UD, opt, tmprval))

    rval = ConvertLongToInt(tmprval)

    Call copy_ibvars
End Function

Function ilbna(ByVal UD As Integer, ByVal udname As String) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilbna = ConvertLongToInt(ibbna32(UD, ByVal udname))

    Call copy_ibvars
End Function

Function ilcac(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilcac = ConvertLongToInt(ibcac32(UD, V))

    Call copy_ibvars
End Function

Function ilclr(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilclr = ConvertLongToInt(ibclr32(UD))

    Call copy_ibvars
End Function

Function ilcmd(ByVal UD As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilcmd = ConvertLongToInt(ibcmd32(UD, ByVal buf, cnt))

    Call copy_ibvars
End Function

Function ilcmda(ByVal UD As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilcmda = ConvertLongToInt(ibcmd32(UD, ByVal buf, cnt))

' When Visual Basic remapping buffer problem is solved, use this:
'    ilcmda = ConvertLongToInt(ibcmda32(ud, ByVal buf, cnt))

    Call copy_ibvars
End Function

Function ilconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilconfig = ConvertLongToInt(ibconfig32(bdid, opt, V))

    Call copy_ibvars
End Function

Function ildev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ildev = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

    Call copy_ibvars
End Function

Function ildma(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ildma = ConvertLongToInt(ibdma32(UD, V))

    Call copy_ibvars
End Function

Function ileos(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ileos = ConvertLongToInt(ibeos32(UD, V))

    Call copy_ibvars
End Function

Function ileot(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ileot = ConvertLongToInt(ibeot32(UD, V))

    Call copy_ibvars
End Function

Function ilfind(ByVal udname As String) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilfind = ConvertLongToInt(ibfind32(ByVal udname))

    Call copy_ibvars
End Function

Function ilgts(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilgts = ConvertLongToInt(ibgts32(UD, V))

    Call copy_ibvars
End Function

Function ilist(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilist = ConvertLongToInt(ibist32(UD, V))

    Call copy_ibvars
End Function

Function illck(ByVal UD As Integer, ByVal V As Integer, ByVal LockWaitTime As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    illck = ConvertLongToInt(iblck32(UD, V, LockWaitTime, ByVal 0))

    Call copy_ibvars
End Function

Function illines(ByVal UD As Integer, lines As Integer) As Integer
    Dim tmplines As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    illines = ConvertLongToInt(iblines32(UD, tmplines))

    lines = ConvertLongToInt(tmplines)

    Call copy_ibvars
End Function

Function illn(ByVal UD As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer) As Integer
    Dim tmpln As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    illn = ConvertLongToInt(ibln32(UD, pad, sad, tmpln))

    ln = ConvertLongToInt(tmpln)

    Call copy_ibvars
End Function

Function illoc(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    illoc = ConvertLongToInt(ibloc32(UD))

    Call copy_ibvars
End Function

Function ilonl(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilonl = ConvertLongToInt(ibonl32(UD, V))

    Call copy_ibvars
End Function

Function ilpad(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilpad = ConvertLongToInt(ibpad32(UD, V))

    Call copy_ibvars
End Function

Function ilpct(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilpct = ConvertLongToInt(ibpct32(UD))

    Call copy_ibvars
End Function

Function ilppc(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilppc = ConvertLongToInt(ibppc32(UD, V))

    Call copy_ibvars
End Function

Function ilrd(ByVal UD As Integer, buf As String, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilrd = ConvertLongToInt(ibrd32(UD, ByVal buf, cnt))

    Call copy_ibvars
End Function

Function ilrda(ByVal UD As Integer, buf As String, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilrda = ConvertLongToInt(ibrd32(UD, ByVal buf, cnt))

' When Visual Basic remapping buffer problem solved, use this:
'    ilrda = ConvertLongToInt(ibrda32(ud, ByVal buf, cnt))

    Call copy_ibvars
End Function

Function ilrdf(ByVal UD As Integer, ByVal filename As String) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilrdf = ConvertLongToInt(ibrdf32(UD, ByVal filename))

    Call copy_ibvars
End Function

Function ilrdi(ByVal UD As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilrdi = ConvertLongToInt(ibrd32(UD, ibuf(0), cnt))

    Call copy_ibvars
End Function

Function ilrdia(ByVal UD As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilrdia = ConvertLongToInt(ibrd32(UD, ibuf(0), cnt))

' When Visual Basic remapping buffer problem solved, use this:
'    ilrdia = ConvertLongToInt(ibrda32(ud, ibuf(0), cnt))

    Call copy_ibvars
End Function

Function ilrpp(ByVal UD As Integer, ppr As Integer) As Integer
    Static tmp_str As String * 2

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilrpp = ConvertLongToInt(ibrpp32(UD, ByVal tmp_str))

    ppr = Asc(tmp_str)

    Call copy_ibvars
End Function

Function ilrsc(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

'  Call the 32-bit DLL.
    ilrsc = ConvertLongToInt(ibrsc32(UD, V))

    Call copy_ibvars
End Function

Function ilrsp(ByVal UD As Integer, spr As Integer) As Integer
    Static tmp_str As String * 2

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL
    ilrsp = ConvertLongToInt(ibrsp32(UD, ByVal tmp_str))

    spr = Asc(tmp_str)

    Call copy_ibvars
End Function

Function ilrsv(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilrsv = ConvertLongToInt(ibrsv32(UD, V))

    Call copy_ibvars
End Function

Function ilsad(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

'  Call the 32-bit DLL.
    ilsad = ConvertLongToInt(ibsad32(UD, V))

    Call copy_ibvars
End Function

Function ilsic(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

'  Call the 32-bit DLL.
    ilsic = ConvertLongToInt(ibsic32(UD))

    Call copy_ibvars
End Function

Function ilsre(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

'  Call the 32-bit DLL.
    ilsre = ConvertLongToInt(ibsre32(UD, V))

    Call copy_ibvars
End Function

Function ilstop(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

'  Call the 32-bit DLL.
    ilstop = ConvertLongToInt(ibstop32(UD))

    Call copy_ibvars
End Function

Function iltmo(ByVal UD As Integer, ByVal V As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

'  Call the 32-bit DLL.
    iltmo = ConvertLongToInt(ibtmo32(UD, V))

    Call copy_ibvars
End Function

Function iltrg(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call 32-bit DLL.
    iltrg = ConvertLongToInt(ibtrg32(UD))

    Call copy_ibvars
End Function

Function ilwait(ByVal UD As Integer, ByVal mask As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilwait = ConvertLongToInt(ibwait32(UD, mask))

    Call copy_ibvars
End Function

Function ilwrt(ByVal UD As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilwrt = ConvertLongToInt(ibwrt32(UD, ByVal buf, cnt))

    Call copy_ibvars
End Function

Function ilwrta(ByVal UD As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilwrta = ConvertLongToInt(ibwrt32(UD, ByVal buf, cnt))

' When the Visual Basic remapping solved, use this:
'    ilwrta = ConvertLongToInt(ibwrta32(ud, ByVal buf, cnt))

    Call copy_ibvars

End Function

Function ilwrtf(ByVal UD As Integer, ByVal filename As String) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilwrtf = ConvertLongToInt(ibwrtf32(UD, ByVal filename))

    Call copy_ibvars
End Function

Function ilwrti(ByVal UD As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilwrti = ConvertLongToInt(ibwrt32(UD, ibuf(0), cnt))

    Call copy_ibvars
End Function

Function ilwrtia(ByVal UD As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilwrtia = ConvertLongToInt(ibwrt32(UD, ibuf(0), cnt))

' When Visual Basic remapping buffer problem solved, use this:
'    ilwrtia = ConvertLongToInt(ibwrta32(ud, ibuf(0), cnt))

    Call copy_ibvars
End Function

Sub PassControl(ByVal boardID As Integer, ByVal addr As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call PassControl32(boardID, addr)

    Call copy_ibvars
End Sub

Sub Ppoll(ByVal boardID As Integer, result As Integer)
    Dim tmpresult As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call PPoll32(boardID, tmpresult)

    result = ConvertLongToInt(tmpresult)

    Call copy_ibvars
End Sub

Sub PpollConfig(ByVal boardID As Integer, ByVal addr As Integer, ByVal lline As Integer, ByVal sense As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call PPollConfig32(boardID, addr, lline, sense)

    Call copy_ibvars
End Sub

Sub PpollUnconfig(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call PPollUnconfig32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub RcvRespMsg(ByVal boardID As Integer, buf As String, ByVal term As Integer)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call RcvRespMsg32(boardID, ByVal buf, cnt, term)

    Call copy_ibvars
End Sub

Sub ReadStatusByte(ByVal boardID As Integer, ByVal addr As Integer, result As Integer)
    Dim tmpresult As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ReadStatusByte32(boardID, addr, tmpresult)

    result = ConvertLongToInt(tmpresult)

    Call copy_ibvars
End Sub

Sub Receive(ByVal boardID As Integer, ByVal addr As Integer, buf As String, ByVal term As Integer)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call Receive32(boardID, addr, ByVal buf, cnt, term)

    Call copy_ibvars
End Sub

Sub ReceiveSetup(ByVal boardID As Integer, ByVal addr As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ReceiveSetup32(boardID, addr)

    Call copy_ibvars
End Sub

Sub ResetSys(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ResetSys32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub Send(ByVal boardID As Integer, ByVal addr As Integer, ByVal buf As String, ByVal term As Integer)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call Send32(boardID, addr, ByVal buf, cnt, term)

    Call copy_ibvars
End Sub

Sub SendCmds(ByVal boardID As Integer, ByVal cmdbuf As String)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(cmdbuf))

' Call the 32-bit DLL.
    Call SendCmds32(boardID, ByVal cmdbuf, cnt)

    Call copy_ibvars
End Sub

Sub SendDataBytes(ByVal boardID As Integer, ByVal buf As String, ByVal term As Integer)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call SendDataBytes32(boardID, ByVal buf, cnt, term)

    Call copy_ibvars
End Sub

Sub SendIFC(ByVal boardID As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call SendIFC32(boardID)

    Call copy_ibvars
End Sub

Sub SendList(ByVal boardID As Integer, addr() As Integer, ByVal buf As String, ByVal term As Integer)
    Dim cnt As Long

' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

    cnt = CLng(Len(buf))

' Call the 32-bit DLL.
    Call SendList32(boardID, addr(0), ByVal buf, cnt, term)

    Call copy_ibvars
End Sub

Sub SendLLO(ByVal boardID As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call SendLLO32(boardID)

    Call copy_ibvars
End Sub

Sub SendSetup(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call SendSetup32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub SetRWLS(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call SetRWLS32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub TestSRQ(ByVal boardID As Integer, result As Integer)
    Call ibwait(boardID, 0)

    If ibsta And &H1000 Then
        result = 1
    Else
        result = 0
    End If

End Sub

Sub TestSys(ByVal boardID As Integer, addrs() As Integer, results() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call TestSys32(boardID, addrs(0), results(0))

    Call copy_ibvars
End Sub

Sub Trigger(ByVal boardID As Integer, ByVal addr As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call Trigger32(boardID, addr)

    Call copy_ibvars
End Sub

Sub TriggerList(ByVal boardID As Integer, addrs() As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call TriggerList32(boardID, addrs(0))

    Call copy_ibvars
End Sub

Sub WaitSRQ(ByVal boardID As Integer, result As Integer)
    Call ibwait(boardID, &H5000)

    If ibsta And &H1000 Then
        result = 1
    Else
        result = 0
    End If
End Sub


Private Function ConvertLongToInt(LongNumb As Long) As Integer

  If (LongNumb And &H8000&) = 0 Then
      ConvertLongToInt = LongNumb And &HFFFF&
  Else
    ConvertLongToInt = &H8000 Or (LongNumb And &H7FFF&)
  End If

End Function

Public Sub RegisterGPIBGlobals()
    Dim rc As Long

    rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
    If (rc = 0) Then
      GPIBglobalsRegistered = 1
    ElseIf (rc = 1) Then
      rc = UnregisterGpibGlobalsForThread
      rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
      GPIBglobalsRegistered = 1
    ElseIf (rc = 2) Then
      rc = UnregisterGpibGlobalsForThread
      ibsta = &H8000
      iberr = EDVR
      ibcntl = &HDEAD37F0
    ElseIf (rc = 3) Then
      rc = UnregisterGpibGlobalsForThread
      ibsta = &H8000
      iberr = EDVR
      ibcntl = &HDEAD37F0
    Else
      ibsta = &H8000
      iberr = EDVR
      ibcntl = &HDEAD37F0
    End If
End Sub

Public Sub UnregisterGPIBGlobals()
    Dim rc As Long

    rc = UnregisterGpibGlobalsForThread
    GPIBglobalsRegistered = 0

End Sub



Public Function ThreadIbsta() As Integer
' Call the 32-bit DLL.
    ThreadIbsta = ConvertLongToInt(ThreadIbsta32())
End Function

Public Function ThreadIberr() As Integer
' Call the 32-bit DLL.
    ThreadIberr = ConvertLongToInt(ThreadIberr32())
End Function

Public Function ThreadIbcnt() As Integer
' Call the 32-bit DLL.
    ThreadIbcnt = ConvertLongToInt(ThreadIbcnt32())
End Function

Public Function ThreadIbcntl() As Long
' Call the 32-bit DLL.
    ThreadIbcntl = ThreadIbcntl32()
End Function

Public Function illock(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    illock = ConvertLongToInt(iblock32(UD))

    Call copy_ibvars
End Function

Public Function ilunlock(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilunlock = ConvertLongToInt(ibunlock32(UD))

    Call copy_ibvars
End Function

Public Sub iblock(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call iblock32(UD)

    Call copy_ibvars
End Sub

Public Sub ibunlock(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibunlock32(UD)

    Call copy_ibvars
End Sub

Public Function illockx(ByVal UD As Integer, ByVal LockWaitTime As Integer, ByVal buf As String) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    illockx = ConvertLongToInt(iblockx32(UD, LockWaitTime, buf))

    Call copy_ibvars
End Function

Public Function ilunlockx(ByVal UD As Integer) As Integer
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    ilunlockx = ConvertLongToInt(ibunlockx32(UD))

    Call copy_ibvars
End Function

Public Sub iblockx(ByVal UD As Integer, ByVal LockWaitTime As Integer, ByVal buf As String)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call iblockx32(UD, LockWaitTime, buf)

    Call copy_ibvars
End Sub

Public Sub ibunlockx(ByVal UD As Integer)
' Check to see if GPIB Global variables are registered
    If (GPIBglobalsRegistered = 0) Then
      Call RegisterGPIBGlobals
    End If

' Call the 32-bit DLL.
    Call ibunlockx32(UD)

    Call copy_ibvars
End Sub