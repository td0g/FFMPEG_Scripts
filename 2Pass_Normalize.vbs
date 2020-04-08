'FFMPEG Two-Pass Audio Normalizing
'Written by Tyler Gerritsen 2020-04-08

'Automates the two-pass audio normalizing filter built in to FFMPEG
'First pass measures the audio levels of the video
'Second pass applies the filter to normalize the audio levels

'TO USE:
'	Drag and drop the video(s) onto the vbscript file

'###################################################################################

			'Script Configuration

'###################################################################################

			'Target audio level
			audioLevel = 2.0

'###################################################################################

			'Changelog

'###################################################################################


'0.1
'	2020-04-08
'	Functional
	

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
				cmdString = chr(34) & binLoc & "ffmpeg.exe" & chr(34)
			elseif fso.FileExists(mainLoc & "ffmpeg.exe") then
				cmdString = chr(34) & mainLoc & "ffmpeg.exe" & chr(34)
			else
				cmdString = "ffmpeg"
			end if
			cmdString = cmdString & " -y -i " & chr(34) & vidName & chr(34)
		Next 
		
'Read video file and get audio levels to a text file
		cmdstring = replace(cmdstring, "/", "\")
		wsh.run "cmd.exe /c " & chr(34) & cmdString &" -af loudnorm=print_format=json -f null - > " & mainLoc & "output.txt 2>&1"	& chr(34), 1, true
		
'Parse output
		Dim iFile
		set iFile = fso.OpenTextFile(mainLoc & "output.txt")
		inputI = -1
		inputTP = 1
		inputLRA = 1
		inputTHRESH = -1
		do until iFile.AtEndOfStream
			line = iFile.ReadLine
			dim objRegEx
			Set objRegEx = CreateObject("VBScript.RegExp")
			objRegEx.Global = True
			objRegEx.Pattern = "[^-0123456789.]"
			trimmedLine = objRegEx.Replace(line, "")
			if instr(line, "input_i") > 0 then
				inputI = CDbl(trimmedLine)
			elseif instr(line, "input_tp") > 0 then
				inputTP=CDbl(trimmedLine)
			elseif instr(line, "input_lra") > 0 then
				inputLRA = CDbl(trimmedLine)
			elseif instr(line, "input_thresh") > 0 then
				inputTHRESH = CDbl(trimmedLine)
			end if
		loop
		iFile.close
		fso.deleteFile mainLoc & "output.txt"
		
'Apply audio normalization
		cmdString = cmdString & " -af loudnorm=linear=true:measured_I="&inputI&":measured_LRA="&inputLRA&":measured_tp="&inputTP&":measured_thresh="&inputTHRESH
		if audioLevel <> 1.0 then cmdString = cmdString & ",volume="&audioLevel
		wsh.run  cmdString & " " & chr(34) & vidNameOut & chr(34)
	end if
end if

