'PCCrash Video ReEncoder
'Written by Tyler Gerritsen 2018-11-27

'PC Crash uses a very inefficient video codec
'This script automates re-encoding the video using a much more efficient encoder

'TO USE:
'	Drag and drop the video(s) (ONE AT A TIME) onto the vbscript file

'FFMPEG is used to re-encode the video using libx264 codec
'The video resolution can also be scaled

'###################################################################################

			'Script Configuration

'###################################################################################

			'Constant Rate Factor - 18 or 19 is DVD quality, 32 is highly compressed
			CRF = 28

			'Resolution - set to 0 (no resize), 720, or 1080
			videoResolution = 0
			
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
	

'###################################################################################

			'Script

'###################################################################################

mainLoc = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")) 'Location of ffmpeg.exe - should be in same folder or \bin subfolder
binLoc = mainLoc & "bin\"	
Set fso = CreateObject("Scripting.FileSystemObject")
if not fso.FileExists(binLoc) and not fso.FileExists(mainLoc) then
	msgbox "FFMPEG not found!" & vbNewLine & vbNewLine & "Please put FFMPEG.exe in same folder as ReEncode script" & vbnewline & "FFMPEG is available from https://www.ffmpeg.org/"
else
	If WScript.Arguments.Count > 0 Then
		For i = 0 to Wscript.Arguments.Count - 1
			vidName = Replace(Wscript.Arguments(i), "\", "/")
			vidNameOut = Replace(vidName, ".", " - REENCODE.")
			if fso.FileExists(binLoc) then
				cmdString = chr(34) & binLoc & "ffmpeg.exe" & chr(34) & " -y -i " & chr(34) & vidName & chr(34)
			else
				cmdString = chr(34) & mainLoc & "ffmpeg.exe" & chr(34) & " -y -i " & chr(34) & vidName & chr(34)
			end if
			if videoResolution > 0 then cmdString = cmdString & " -vf " & chr(34) & "scale=-2:" & videoResolution & chr(34)
			cmdString = cmdString & " -vcodec libx264 -crf " & CRF & " -profile:v baseline -level 3.0 -pix_fmt yuv420p -movflags faststart " & params & " " & chr(34) & vidNameOut & chr(34) & vbNewLine
		Next 
		'msgbox cmdstring
		dim wSh														'Prepare windows shell object
		Set wSh = WScript.CreateObject("WScript.Shell")
		wsh.run cmdString, 1, TRUE									'Execute
	end if
end if
