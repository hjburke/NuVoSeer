' =============================================================================
'
' Control Script for NuVo Concerto Whole House Audio System
'
' Rev 20110609
'
' =============================================================================
'
' Change History
'
'		20110609	HJB		- Initial Release by Howard Burke
' =============================================================================

' =============================================================================
' S c r i p t   V a r i a b l e s   a n d   C o n s t a n t s
' =============================================================================

Const ScriptName = "NuVo.vb"

Const HouseCode As String = "Y"

Const SerialPort = 3

Const NuvoZoneCodeStart As Integer = 1		' The Device Code (DC) that will hold the first Nuvo Zone device
Const NuvoZoneSourceStart As Integer = 9	' The Device Code (DC) that will hold the first Nuvo Zone Source device
Const NuvoZoneVolumeStart As Integer = 17	' The Device Code (DC) that will hold the first Nuvo Zone Volume device
Const NuvoZones As Integer = 8				' The number of Nuvo Zones
Const NuvoSources As Integer = 6			' The number of Nuvo Sources

Const NuvoiTunesSource As Integer = 1		' The source used for the iTunes player

' ============================================================================
'
' HomeSeer Log Output
'
' Actual errors are always logged to the HomeSeer log
'
' For problems, the following constant, scTraceLevel, can be set as follows:
'
'	0 = No trace output
'	1 = Minimal tracing (main methods)
'	2 = Additional method tracing and other activity
'	3 = All methods traced, variable tracing plus additional diagnostic output
'	4 = Very low level tracing (lots and lots of output)
'
Const TraceLevel As Integer = 2

' =============================================================================
'
' Main()
'
' =============================================================================

Public Sub Main(ByVal Parms As Object)
	OpenComPort()
    hs.RegisterStatusChangeCB(ScriptName, "StatusChangeCallBack")	
End Sub

'Public Class hsclass
'Public Sub New(ByVal Parms As Object)
'	TraceIt(2, "Entering RegStatusCB")
'	hs.RegisterEventCB(1024,me)
'End Sub
'
'Public Sub HSEvent(ByVal Parms As Object)
'    Dim hc As String
'    Dim dc As Integer
'	Dim newval As Integer
'	Dim oldval As Integer
'
'	Select Case Parms(0)
'		Case 1024
'		dc = Parms(1)
'		hc = Parms(2)
'		newval = Parms(3)
'		oldval = Parms(4)
'		
'		TraceIt(2, "Got HSEvent dc=" & CStr(dc) & " hc=" & CStr(hc) & " newval=" & CStr(newval) & " oldval=" & CStr(oldval))
'		
'	End Select
'
'End Sub
'End Class

' =============================================================================
'
' ResetComPort()
'
' =============================================================================

Public Sub ResetComPort(ByVal Parms As Object)
	CloseComPort()
	hs.WaitSecs(1)	
	OpenComPort()
End Sub

' =============================================================================
'
' InstallDevices()
'
' Create all the devices
'
' =============================================================================

Public Sub InstallDevices(ByVal Dummy As Object)

	Dim zone As Integer
	Dim dv As Scheduler.Classes.DeviceClass

	For zone = 1 To NuvoZones	
'		dv = hs.NewDeviceEx("Nuvo Zone " & zone.ToString)
'		dv.hc = HouseCode
'		dv.dc = zone
'		dv.dev_type_string = "Nuvo Keypad"	
		
'		dv = hs.NewDeviceEx("Nuvo Zone " & zone.ToString & " Source")
'		dv.hc = HouseCode
'		dv.dc = NuvoZoneSourceStart+zone-1
'		dv.dev_type_string = "Nuvo Keypad"	
'		hs.DeviceValuesAdd(dv.hc & dv.dc, _
'			"Source 1" & chr(2) & "1" & chr(1) & _
'			"Source 2" & chr(2) & "2" & chr(1) & _
'			"Source 3" & chr(2) & "3" & chr(1) & _
'			"Source 4" & chr(2) & "4" & chr(1) & _
'			"Source 5" & chr(2) & "5" & chr(1) & _
'			"Source 6" & chr(2) & "6", True)
			
'		dv = hs.NewDeviceEx("Nuvo Zone " & zone.ToString & " Volume")
'		dv.hc = HouseCode
'		dv.dc = NuvoZoneVolumeStart+zone-1
'		dv.dev_type_string = "Nuvo Keypad"	
'		'hs.DeviceButtonAdd(dv.hc & dv.dc, ScriptName(""ZoneMute"",zone),"Mute")
'		'hs.DeviceButtonAdd(dv.hc & dv.dc, ScriptName(""ZoneUnMute"",zone),"UnMute")
	Next

end sub

' =============================================================================
'
' StatusChangeCallBack()
'
' HomeSeer will call us any time a device changes state.  We use this to track when the user wants to control a zone
'
' =============================================================================
Public Sub StatusChangeCallBack(ByVal parms)

    Dim hc As String
    Dim dc As Integer
    Dim status As Integer
	Dim zone As Integer

    Try
        ' Get the info on the device that changed
        hc = parms(0)
        dc = parms(1)
        status = parms(2)

        ' Eliminate all devices that are not ours
        If hc <> HouseCode Then Return

		
        ' Handle any changes to the Nuvo zones
        If dc >= NuvoZoneCodeStart And dc <= NuvoZoneCodeStart + NuvoZones Then
			zone = dc
            ' Send the command to control the zone
            If status = 2 Then ' Turn On 
				hs.WriteLog(ScriptName, "Turn ON NuVo Zone " & zone.ToString)
				TurnOn(zone.ToString)
            ElseIf status = 3 Then ' Turn Off
				hs.WriteLog(ScriptName, "Turn OFF NuVo Zone " & zone.ToString)
				TurnOff(zone.ToString)
            End If
        ' Handle any source changes
'        ElseIf dc >= NuvoZoneSourceStart And dc <= NuvoZoneSourceStart + NuvoZones Then
'			zone = dc-NuvoZoneSourceStart
'			hs.WriteLog(ScriptName, "Set NuVo Zone " & zone.ToString & " to Source " & status.ToString)
'			SetSource(zone,status)
        End If

    Catch ex As Exception
        hs.WriteLog(ScriptName, "Error in StatusChangeCallback" & ex.ToString)
    End Try

End Sub

' =============================================================================
'
' AllOff(zone)
'
' Turn off power to all zones
'
' =============================================================================
Sub AllOff(ByVal dummy As Object)

	TraceIt(2, "Entering AllOff")

	SendCommand("ALLOFF")
	
	TraceIt(2, "Exiting AllOff")
	
End Sub

' =============================================================================
'
' TurnOn(zone)
'
' Turn on power for a zone
'
'	zone = 1 to NuvoZones
'
' =============================================================================
Sub TurnOn(ByVal zonestr As Object)

	Dim zone As Integer
	zone = CInt(zonestr)

	TraceIt(2, "Entering TurnOn(" & CStr(zone) & ")")

	If (zone < 1) Or (zone > NuvoZones) Then	
		LogIt("TurnOn: Zone " & zone & " is out of range")
		TraceIt(2, "Exiting TurnOn(" & CStr(zone) & ")")
		Exit Sub		
	End If
	
	SendCommand("Z" & zone & "ON")

	TraceIt(2, "Exiting TurnOn(" & CStr(zone) & ")")
	
End Sub

' =============================================================================
'
' TurnOff(zone)
'
' Turn off power for a zone
'
'	zone = 1 to NuvoZones
'
' =============================================================================
Sub TurnOff(ByVal zonestr As Object)

	Dim zone As Integer
	zone = CInt(zonestr)

	TraceIt(2, "Entering TurnOff(" & CStr(zone) & ")")

	If (zone < 1) Or (zone > NuvoZones) Then	
		LogIt("TurnOff: Zone " & zone & " is out of range")
		TraceIt(2, "Exiting TurnOff(" & CStr(zone) & ")")
		Exit Sub		
	End If
	
	SendCommand("Z" & zone & "OFF")
	
	TraceIt(2, "Exiting TurnOff(" & CStr(zone) & ")")

End Sub

' =============================================================================
'
' VolumeUp(zone)
'
' Turn the volume up for a zone zone
'
'	zone = 1 to NuvoZones
'
' NOTE: In order to change volume, the zone must first be on.  If it is not
'       already on, this routine will turn the zone on.
' =============================================================================
Sub VolumeUp(ByVal zonestr As String)

	Dim zone As Integer
	zone = CInt(zonestr)

	TraceIt(2, "Entering VolumeUp(" & CStr(zone) & ")")

	If (zone < 1) Or (zone > NuvoZones) Then	
		LogIt("VolumeUp: Zone " & zone & " is out of range")
		TraceIt(2, "Exiting VolumeUp(" & CStr(zone) & ")")
		Exit Sub		
	End If

	If (hs.IsOff(HouseCode & zone) = true) Then
		TurnOn(CStr(zone))
	End If
	
	SendCommand("Z" & zone & "VOL+")

	TraceIt(2, "Exiting VolumeUp(" & CStr(zone) & ")")

End Sub

' =============================================================================
'
' VolumeDown(zone)
'
' Turn the volume down for a zone zone
'
'	zone = 1 to NuvoZones
'
' NOTE: In order to change volume, the zone must first be on.  If it is not
'       already on, this routine will turn the zone on.
' =============================================================================
Sub VolumeDown(ByVal zonestr As String)

	Dim zone As Integer
	zone = CInt(zonestr)

	TraceIt(2, "Entering VolumeDown(" & CStr(zone) & ")")

	If (zone < 1) Or (zone > NuvoZones) Then	
		LogIt("VolumeDown: Zone " & zone & " is out of range")
		TraceIt(2, "Exiting VolumeDown(" & CStr(zone) & ")")
		Exit Sub		
	End If

	If (hs.IsOff(HouseCode & zone) = true) Then
		TurnOn(CStr(zone))
	End If
	
	SendCommand("Z" & zone & "VOL-")

	TraceIt(2, "Exiting VolumeDown(" & CStr(zone) & ")")

End Sub

' =============================================================================
'
' SetSource("zone,source")
'
' Set the source for a zone
'
'	zone = 1 to NuvoZones
'	source = 1 to NuvoSources
'
' NOTE: In order to change a source, the zone must first be on.  If it is not
'       already on, this routine will turn the zone on.
' =============================================================================
Sub SetSource(ByVal params As Object)

	Dim zone As Integer, source As Integer

	zone = CInt(hs.stringitem(params,1,","))
	source = CInt(hs.stringitem(params,2,","))
	
	TraceIt(2, "Entering SetSource(" & CStr(zone) & Chr(44) & CStr(source) & ")")

	If (zone < 1) Or (zone > NuvoZones) Then
		LogIt("SetSource: Zone " & zone & " is out of range")
		TraceIt(2, "Exiting SetSource(" & CStr(zone) & Chr(44) & CStr(source) & ")")
		Exit Sub
	End If

	If (source < 1) Or (source > NuvoSources) Then
		LogIt("SetSource: Source " & source & " is out of range")
		TraceIt(2, "Exiting SetSource(" & CStr(zone) & Chr(44) & CStr(source) & ")")
		Exit Sub	
	End If

	If (hs.IsOff(HouseCode & zone) = true) Then
		TurnOn(CStr(zone))
	End If
	
	SendCommand("Z" & zone & "SRC" & source)

	TraceIt(2, "Exiting SetSource(" & CStr(zone) & Chr(44) & CStr(source) & ")")

End Sub

' =============================================================================
'
' SourceDisplay("source,message,perm")
'
' Display a message on the zones on this source
'
'	source = 1 to NuvoSources
'	message = text to be displayed
'   perm = true=display message until another message is displayed
'
' =============================================================================
Sub SourceDisplay(ByVal params As Object)
  
	Dim source As Integer, message As String, perm As Boolean, cmd As String

	source = CInt(hs.stringitem(params,1,","))
	message = hs.stringitem(params,2,",")
	perm = CBool(hs.stringitem(params,3,","))

	TraceIt(2, "Entering SourceDisplay(" & CStr(source) & Chr(44) & message & ")")
	
	If (perm=true) Then
		cmd="STR+"
	Else
		cmd="STR"
	End If	

	If (source < 1) Or (source > NuvoSources) Then
		LogIt("SourceDisplay: Source " & source & " is out of range")
		TraceIt(2, "Exiting SourceDisplay(" & CStr(source) & Chr(44) & message & ")")
		Exit Sub
	End If

	SendCommand("S" & source & cmd & Chr(34) & message & Chr(34))

	TraceIt(2, "Exiting SourceDisplay(" & CStr(source) & Chr(44) & message & ")")

End Sub

' =============================================================================
'
' SourceDisplayNowPlaying("source,perm")
'
' Display a message on the zones on this source
'
'	source = 1 to NuvoSources
'   perm = true=display message until another message is displayed
'
' =============================================================================
Sub SourceDisplayNowPlaying(ByVal params As Object)
  
	Dim source As Integer, message As String, perm As Boolean, postfix As String, cmd As String

	source = CInt(hs.stringitem(params,1,","))
	perm = CBool(hs.stringitem(params,2,","))
	postfix = hs.stringitem(params,3,",")

	' We have to wait 1 second, because if you call this from an iTunes trigger,
	' it fires before the Current info is updated, so you get the details for the
	' last track playing, not the current one.
	hs.WaitSecs(1)
	
	message = (hs.Plugin("iTunes").MusicAPI.CurrentTrack + " - " + hs.Plugin("iTunes").MusicAPI.CurrentArtist + postfix).ToUpper
	
	TraceIt(2, "Entering SourceDisplayNowPlaying(" & CStr(source) & Chr(44) & message & ")")
	
	If (perm=true) Then
		cmd="STR+"
	Else
		cmd="STR"
	End If	

	If (source < 1) Or (source > NuvoSources) Then
		LogIt("SourceDisplayNowPlaying: Source " & source & " is out of range")
		TraceIt(2, "Exiting SourceDisplayNowPlaying(" & CStr(source) & Chr(44) & message & ")")
		Exit Sub
	End If

	SendCommand("S" & source & cmd & Chr(34) & message & Chr(34))

	TraceIt(2, "Exiting SourceDisplayNowPlaying(" & CStr(source) & Chr(44) & message & ")")

End Sub

' =============================================================================
'
' ZoneDisplay("zone,message,perm")
'
' Display a message on the zone
'
'	zone = 1 to NuvoZones
'	message = text to be displayed
'   perm = true=display message until another message is displayed
'
' =============================================================================
Sub ZoneDisplay(ByVal params As Object)
  
	Dim zone As Integer, message As String, perm As Boolean, cmd As String

	zone = CInt(hs.stringitem(params,1,","))
	message = hs.stringitem(params,2,",")
	perm = CBool(hs.stringitem(params,3,","))

	TraceIt(2, "Entering ZoneDisplay(" & CStr(zone) & Chr(44) & message & ")")
	
	If (perm=true) Then
		cmd="STR+"
	Else
		cmd="STR"
	End If	

	If (zone < 1) Or (zone > NuvoZones) Then
		LogIt("ZoneDisplay: Zone " & zone & " is out of range")
		TraceIt(2, "Exiting ZoneDisplay(" & CStr(zone) & Chr(44) & message & ")")
		Exit Sub
	End If

	message = hs.DeviceString("^1")
	
	SendCommand("Z" & zone & cmd & Chr(34) & message & Chr(34))

	TraceIt(2, "Exiting ZoneDisplay(" & CStr(zone) & Chr(44) & message & ")")

End Sub

' =============================================================================
' C o m m u n i c a t i o n s   M e t h o d s
' =============================================================================

' =============================================================================
'
' OpenComPort
'
' Open communications to the device
'
' =============================================================================
Private Sub OpenComPort()

	TraceIt(2, "Entering OpenComPort(" & CStr(SerialPort) & ")")

	Dim	err
	
	err=hs.OpenComPort(SerialPort,"9600,n,8,1",1,ScriptName,"ProcessResponse", vbCr)
        
	If err <> "" Then
		 LogIt("Error opening COM" & SerialPort & ": " & err)
	Else
	     LogIt("COM" & SerialPort & " Setup complete")
    End If 

	TraceIt(2, "Exiting OpenComPort(" & CStr(SerialPort) & ")")
	
End Sub

' =============================================================================
'
' CloseComPort
'
' Close communications to the device
'
' =============================================================================
Private Sub CloseComPort()

	TraceIt(2, "Entering CloseComPort(" & CStr(SerialPort) & ")")

	hs.CloseComPort(SerialPort)

	TraceIt(2, "Exiting CloseComPort(" & CStr(SerialPort) & ")")

End Sub

' =============================================================================
'
' SendCommand
'
' Send "command" to the device at Port "port"
'
' =============================================================================
Private Sub SendCommand(ByVal command As String)
	'Dim pObj As Object = hs.Plugin("Global Cache")

	TraceIt(2, "Entering SendCommand(" & command & ")")
	
	' Send our command
	hs.SendToComPort(SerialPort, "*" & command & vbCr)
	'pObj.SendSerialData(1,2,"*" & command & vbCr)
	
	TraceIt(2, "Exiting SendCommand(" & command & ")")

End Sub

' =============================================================================
' ProcessResponse(data)
'
' Process Responses from the device
'
' =============================================================================
Public Sub ProcessResponse(data)

	TraceIt(1, "ProcessResponse(" & data & ")")
	
	TraceIt(2, "Entering ProcessResponse(" & data & ")")

	Dim nPos As Integer		'position of "#" within response string
	Dim value				'value of the response

	Dim zone As Integer		' Nuvo Zone
	Dim source As Integer	' Nuvo Source
	Dim button As String	' Nuvo Button Pressed

	Dim paramArr() As String
	Dim count As Integer
	Dim dv As Scheduler.Classes.DeviceClass

	TraceIt(3, "Length of data = " & CStr(Len(data)))

	' Trim everything up to and including the "#" off the front of Response
	nPos = InStr(data, "#")
	TraceIt(3, "nPos = " & CStr(nPos))
	If nPos > 0 Then
		TraceIt(3, "Trimming up to # from front of Response")
		data = Right(data, Len(data) - nPos)
	End If 

	' Trim <Cr> off the end of the response
	If Right(data,2) = vbCr Then
		TraceIt(3, "Trimming vbCr from end of Response")
		data = Left(data, Len(data)-2)
	End If 
		
	If Left(data,1) = "Z" And InStr(data, ",") Then
		' =========================
		'  Zx, - Zone Status Info
		' =========================
		
		TraceIt(2, "Got Zx Response: " & CStr(data))

		TraceIt(3, "Response Length: " & Len(data))
		
		zone = CInt(data.Substring(1,2))
		TraceIt(3, "Zone: " & zone)
		
		value = Right(data, Len(data)-3)
		paramArr = value.split(",")
		For count = 0 To paramArr.Length - 1
			If Left(paramArr(count),5) = "PWRON" Then
				hs.SetDeviceStatus(HouseCode & CStr(zone),2)
				TraceIt(2, "Zone " & zone & " is ON")
			ElseIf Left(paramArr(count),6) = "PWROFF" Then
				hs.SetDeviceStatus(HouseCode & CStr(zone),3)
				TraceIt(2, "Zone " & zone & " is OFF")				
			ElseIf Left(paramArr(count),3) = "SRC" Then
				source = CInt(paramArr(count).Substring(3,1))
				hs.SetDeviceValue(HouseCode & CStr(zone+NuvoZoneSourceStart-1),source)
				TraceIt(2, "Source: " & source.ToString)
			ElseIf Left(paramArr(count),3) = "VOL" Then
				TraceIt(2, "Vol: " & paramArr(count))
			End If			
		Next
	ElseIf Left(data,1) = "S" And data.Substring(2,1) = "Z" Then
		' =========================
		'  Sx - Source Status Info
		' =========================
		
		TraceIt(2, "Got Sx Response: " & CStr(data))
		
		source = CInt(data.Substring(1,1))
		zone = CInt(data.Substring(3,2))
		button = data.Substring(5)

		TraceIt(3, "Source: " & Cstr(source))
		TraceIt(3, "Zone: " & CStr(zone))
		TraceIt(3, "Button: " & button)
		
		If source = NuvoiTunesSource Then
			Select Case button
				Case "FWD"
					TraceIt(2, "iTunes Next Track")
					hs.Plugin("iTunes").MusicAPI.TrackNext()
				Case "RWD"
					TraceIt(2, "iTunes Previous Track")
					hs.Plugin("iTunes").MusicAPI.TrackPrev()
				Case "PAUSE"
					TraceIt(2, "iTunes Pause")
					If hs.Plugin("iTunes").MusicAPI.PlayerState() = 1 Then
						hs.Plugin("iTunes").MusicAPI.PauseIfPlaying()
						SourceDisplayNowPlaying(CStr(NuvoiTunesSource)+",true, [PAUSED]")
					End If
				Case "PLAY"
					TraceIt(2, "iTunes Play")
					If hs.Plugin("iTunes").MusicAPI.PlayerState() = 3 Then
						hs.Plugin("iTunes").MusicAPI.PlayIfPaused()
					Else
						hs.Plugin("iTunes").MusicAPI.Play()
					End If
					SourceDisplayNowPlaying(CStr(NuvoiTunesSource)+",true")
				Case "STOP"
					TraceIt(2, "iTunes Stop")
					hs.Plugin("iTunes").MusicAPI.StopPlay()							
					SourceDisplay(CStr(NuvoiTunesSource)+",ITUNES STOPPED,true")
				Case Else
					TraceIt(2, "Unhandled Button Press " & button)
			End Select
		End If
	
	ElseIf data = "AllOff" Then
		' =========================
		'  All Zones Are Off
		' =========================

		TraceIt(2, "Got All Zones Off response: " & CStr(data))
		
		For zone = 1 To NuvoZones
			hs.SetDeviceStatus(HouseCode & CStr(zone),3)	
		Next
	End If 

	TraceIt(2, "Exiting ProcessResponse(" & data & ")")

End Sub


' =============================================================================
' L o g g i n g   a n d   T r a c i n g   M e t h o d s
' =============================================================================
' =============================================================================
'
' LogIt
'
' Unconditionally log a message to the HomeSeer log
'
' =============================================================================
Private Sub LogIt(ByVal message As String)

	hs.WriteLog(ScriptName, message)

End Sub
 
' =============================================================================
'
' TraceIt
'
' Output a trace message to the HomeSeer log
'
' =============================================================================
Private Sub TraceIt(ByVal level As Integer, ByVal message As String)

	If (level <= TraceLevel) Then
		hs.WriteLog(ScriptName, "(" & CStr(level) & ") " & message)
	End If

End Sub