'PCCrash Video ReEncoder TEMPLATE
'Written by Tyler Gerritsen 2018-11-27

'PC Crash uses a very inefficient video codec
'This script automates re-encoding the video using a much more efficient encoder

'TO USE:
'	Drag and drop the video(s) onto the vbscript file

'FFMPEG is used to re-encode the video using libx264 codec

'For compatibility, the recommendations on the ffmpeg website have been followed
'  https://trac.ffmpeg.org/wiki/Encode/H.264#Compatibility

'###################################################################################

			'Script Configuration

'###################################################################################

			'FFMPEG video filters to apply
			'Separate filters with commas
			vfilters = ""
			
			'FFMPEG audio filters to apply
			'Separate filters with commas
			afilters = ""
			
			'Constant Rate Factor - 18 or 19 is DVD quality, 32 is highly compressed
			CRF = 19
			
			'Extra parameters
			params = ""

			
'###################################################################################

			'Changelog

'###################################################################################


'0.1
'	2018-11-27
'	Functional

'0.1b
'	2019-06-21
'	Added -g 20 to make every 20th frame a keyframe

'0.1c
'	2019-10-03
'	If FFMPEG.exe not found then will quit

'0.2
'	2020-04-08
'	Created template from old version
'	Now checks System PATH for ffmpeg.exe
	

'###################################################################################

			'Script

'###################################################################################


mainLoc = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")) 'Location of ffmpeg.exe - should be in same folder or \bin subfolder
binLoc = mainLoc & "bin\"	
PATHloc = binLoc
Set fso = CreateObject("Scripting.FileSystemObject")
dim wSh														'Prepare windows shell object
Set wSh = WScript.CreateObject("WScript.Shell")

'If ffmpeg.exe isn't in same folder of subfolder then check system PATH
if not fso.FileExists(binLoc & "ffmpeg.exe") and not fso.FileExists(mainLoc & "ffmpeg.exe") then
	if instr(lcase(wsh.ExpandEnvironmentStrings( "PATH=%PATH%" )),"ffmpeg") > 0 then
		PATHloc = wsh.ExpandEnvironmentStrings( "PATH=%PATH%" )
		PATHloc = replace(PATHloc, "PATH=","")
		PATHlocArray = split(PATHloc,";")
		for i = 0 to uBound(PATHlocArray)-1	'Loop through PATH for ffmpeg reference
			if instr(lcase(PATHlocArray(i)),"ffmpeg") > 0 then
				PATHloc = PATHlocArray(i)
				if right(PATHloc,1) <> "\" then PATHloc = PATHloc & "\"
				exit for
			end if
		next
	end if
end if

'If ffmpeg.exe isn't found then notify user
if not fso.FileExists(binLoc & "ffmpeg.exe") and not fso.FileExists(mainLoc & "ffmpeg.exe") and not fso.FileExists(PATHloc & "ffmpeg.exe")then
	msgbox "FFMPEG not found!" & vbNewLine & vbNewLine & "Please put FFMPEG.exe in same folder as ReEncode script" & vbnewline & "FFMPEG is available from https://www.ffmpeg.org/"

'Otherwise write and execute command
else
	If WScript.Arguments.Count > 0 Then
		For i = 0 to Wscript.Arguments.Count - 1
			vidName = Replace(Wscript.Arguments(i), "\", "/")
			vidNameOut = Replace(vidName, ".", " - REENCODE.")
			vidNameOut = left(vidNameOut, inStrRev(vidNameOut, ".")) & "mp4"
			if fso.FileExists(binLoc & "ffmpeg.exe") then
				cmdString = chr(34) & binLoc & "ffmpeg.exe" & chr(34) & " -y -i " & chr(34) & vidName & chr(34)
			elseif fso.FileExists(mainLoc & "ffmpeg.exe") then
				cmdString = chr(34) & mainLoc & "ffmpeg.exe" & chr(34) & " -y -i " & chr(34) & vidName & chr(34)
			else
				cmdString = "ffmpeg" & " -y -i " & chr(34) & vidName & chr(34)
			end if
			if vfilters <> "" then cmdString = cmdString & " -vf " & chr(34) & vfilters & chr(34)
			if afilters <> "" then cmdStirng = cmdString & " -af " & chr(34) & afilters & chr(34)
			cmdString = cmdstring & " -vcodec libx264 -crf " & CRF & " -profile:v baseline -level 3.0 -pix_fmt yuv420p -movflags faststart " & params & " " & chr(34) & vidNameOut & chr(34)
		Next 
		cmdstring = replace(cmdstring, "/", "\")
		wsh.run cmdString									'Execute
	end if
end if

