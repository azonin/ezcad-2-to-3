' Update the parameters below to match the source machine where the file came from, and the target machine for the newly generated file
bln_SourceMOPA = False
bln_TargetMOPA = True

n_SourceWattage = 30
n_TargetWattage = 100

n_SourceLens = 110
n_TargetLens = 110

s_Description = "Converted by ezcad-2-to-3"

WScript.Echo Chr(16) & String(115,Chr(150)) & Chr(17)
WScript.Echo Chr(166) & " Convert EZCAD2 to EZCAD3 Marking Parameters Library Converter v1.21                                               " & Chr(166)
WScript.Echo Chr(166) & " Copyright " & Chr(169) & " 2022 Alex Zonin                                                                                       " & Chr(166)
WScript.Echo Chr(166) & String(115,Chr(150)) & Chr(166)
WScript.Echo Chr(166) & " - Imports the MarkParam.lib file (EZCAD2 format)                                                                  " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - Converts the old EZCAD2 parameter format to the new EZCAD3 format                                               " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - Saves the new file as MarkParamlib.ini-NNNW-NNNmm (appends to the existing file in the same folder)             " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - All files have to be in the same folder as this script                                                          " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - If you want to change your time correction settings (TC values), you can do it in the section-template.txt file " & Chr(166)
WScript.Echo Chr(16) & String(115,Chr(150)) & Chr(17)
WScript.Echo ""

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
s_StartDir = fso.GetAbsolutePathName(".") & "\"

s_OldINIFile = s_StartDir + "MarkParam.lib"
s_NewINIFile = s_StartDir + "MarkParamlib.ini-" & n_TargetWattage & "W-" & n_TargetLens & "mm"
s_SectionTemplateFile = s_StartDir + "section-template.txt"

If fso.FileExists(s_SectionTemplateFile) Then
	Set obj_File = fso.OpenTextFile(s_SectionTemplateFile,1)
			s_SectionTemplate = obj_File.ReadAll
	Set obj_File = Nothing
End If

If fso.FileExists(s_OldINIFile) Then
	WScript.Echo "Processing """ & s_OldINIFile & """"
	Set obj_File = fso.OpenTextFile(s_OldINIFile,1)
		Do While obj_File.AtEndOfStream <> True
			s_TempBuff = Trim(obj_File.ReadLine)

			'Skip the legacy LASERMODE block, not present in the new EZCAD3 file
			If s_TempBuff = "[LASERMODE]" Then
				Do While Len(s_TempBuff) > 1
					s_TempBuff = obj_File.ReadLine
				Loop

			'Read the section name
			ElseIf Left(s_TempBuff,1) = "[" Then
				'Initialize the Target Section
				s_TargetSection = s_SectionTemplate
				
				'Extract the Section Name
				s_SectionName = Trim(Mid(s_TempBuff,2,Len(s_TempBuff)-2))
				
				s_TargetSection = Replace(s_TargetSection,"[SECTION-NAME]","[" & s_SectionName & "]")
				s_TargetSection = Replace(s_TargetSection,"NAME=NAME","NAME=" & s_SectionName)
				s_TargetSection = Replace(s_TargetSection,"DESC=DESC","DESC=" & s_Description)

			ElseIf Left(s_TempBuff,10) = "MARKSPEED=" Then
				s_TargetSection = Replace(s_TargetSection,"MARKSPEED=MARKSPEED","MARKSPEED=" & FormatNumber(CDbl(Trim(Right(s_TempBuff,Len(s_TempBuff)-10))),6,-1,0,0))

			ElseIf Left(s_TempBuff,11) = "POWERRATIO=" Then
				n_SourcePowerRatio = CDbl(Right(s_TempBuff,Len(s_TempBuff)-11))
				n_TargetPowerRatio = Ceiling(n_SourceWattage * (n_SourcePowerRatio * n_TargetLens / n_SourceLens) / n_TargetWattage)

				s_TargetSection = Replace(s_TargetSection,"POWERRATIO=POWERRATIO","POWERRATIO=" & FormatNumber(n_TargetPowerRatio,6,-1,0,0))

			ElseIf Left(s_TempBuff,12) = "QPULSEWIDTH=" Then
				'The default Pulse Width of non-MOPA lasers is 200ns. So we have to ignore what's in the source file, if the target is a MOPA, as that source setting (likely 10ns) is wrong.
				'Non-MOPA machines operate at 200ns, completely ignoring the 10ns setting in the parameter library
				If (bln_SourceMOPA = False) And (bln_TargetMOPA = True) Then
					s_TargetSection = Replace(s_TargetSection,"QPULSEWIDTH=QPULSEWIDTH","QPULSEWIDTH=200.000000")

				'If both Source and Target are MOPA, we have to snap the values to the ones available in the target machine configuration
				'Check your /PARAM/LaserPW.ini and adjust the values in the SnapToMOPA function below!
				ElseIf (bln_SourceMOPA = True) And (bln_TargetMOPA = True) Then
					s_TargetSection = Replace(s_TargetSection,"QPULSEWIDTH=QPULSEWIDTH","QPULSEWIDTH=" & FormatNumber(SnapToMOPA(CDbl(Trim(Right(s_TempBuff,Len(s_TempBuff)-12)))),6,-1,0,0))

				Else
					s_TargetSection = Replace(s_TargetSection,"QPULSEWIDTH=QPULSEWIDTH","QPULSEWIDTH=" & FormatNumber(CDbl(Trim(Right(s_TempBuff,Len(s_TempBuff)-12))),6,-1,0,0))
				End If
			
			ElseIf Left(s_TempBuff,11) = "WOBBLEMODE=" Then
				s_TargetSection = Replace(s_TargetSection,"WOBBLEMODE=WOBBLEMODE","WOBBLEMODE=" & Trim(Right(s_TempBuff,Len(s_TempBuff)-11)))
			
			ElseIf Left(s_TempBuff,15) = "WOBBLEDIAMETER=" Then
				s_TargetSection = Replace(s_TargetSection,"WOBBLEDIAMETER=WOBBLEDIAMETER","WOBBLEDIAMETER=" & FormatNumber(CDbl(Trim(Right(s_TempBuff,Len(s_TempBuff)-15))),6,-1,0,0))
		
			ElseIf Left(s_TempBuff,11) = "WOBBLEDIST=" Then
				s_TargetSection = Replace(s_TargetSection,"WOBBLEDIST=WOBBLEDIST","WOBBLEDIST=" & FormatNumber(CDbl(Trim(Right(s_TempBuff,Len(s_TempBuff)-11))),6,-1,0,0))
	
			ElseIf Left(s_TempBuff,5) = "FREQ=" Then
				s_TargetSection = Replace(s_TargetSection,"FREQF=FREQF","FREQF=" & FormatNumber(CDbl(Trim(Right(s_TempBuff,Len(s_TempBuff)-5))),6,-1,0,0))

			ElseIf Len(s_TempBuff) <= 1 Then
				Set obj_OutputFile = fso.OpenTextFile(s_NewINIFile, 8, True) 'Appends to the existing MarkParamlib.ini if there is one in the same folder as this script
				obj_OutputFile.Write s_TargetSection
				obj_OutputFile.Write vbCrLf
				Set obj_OutputFile = Nothing				
				WScript.Echo "Processed: """ & s_SectionName & """"
			End If

		Loop
	Set obj_File = Nothing
End If

Set fso = Nothing

WScript.Echo "Appended to """ & s_NewINIFile & """"

Function Ceiling(Number)
    Ceiling = Int(Number)
    If Ceiling <> Number Then Ceiling = Ceiling + 1
    If Ceiling > 95 Then Ceiling = 95
End Function

'Check your /PARAM/LaserPW.ini and adjust the values in the SnapToMOPA function below!
Function SnapToMOPA(n_PulseWidth)
	'These values a for JPT M7 MOPA
	If (n_PulseWidth >= 1) And (n_PulseWidth <= 2) Then
		SnapToMOPA = 2
	ElseIf (n_PulseWidth >= 3) And (n_PulseWidth <= 4) Then SnapToMOPA = 4
	ElseIf (n_PulseWidth >= 5) And (n_PulseWidth <= 7) Then SnapToMOPA = 6
	ElseIf (n_PulseWidth >= 8) And (n_PulseWidth <= 10) Then SnapToMOPA = 9
	ElseIf (n_PulseWidth >= 11) And (n_PulseWidth <= 15) Then SnapToMOPA = 13
	ElseIf (n_PulseWidth >= 16) And (n_PulseWidth <= 24) Then SnapToMOPA = 20
	ElseIf (n_PulseWidth >= 25) And (n_PulseWidth <= 37) Then SnapToMOPA = 30
	ElseIf (n_PulseWidth >= 38) And (n_PulseWidth <= 49) Then SnapToMOPA = 45
	ElseIf (n_PulseWidth >= 50) And (n_PulseWidth <= 57) Then SnapToMOPA = 55
	ElseIf (n_PulseWidth >= 58) And (n_PulseWidth <= 69) Then SnapToMOPA = 60
	ElseIf (n_PulseWidth >= 70) And (n_PulseWidth <= 89) Then SnapToMOPA = 80
	ElseIf (n_PulseWidth >= 90) And (n_PulseWidth <= 124) Then SnapToMOPA = 100
	ElseIf (n_PulseWidth >= 125) And (n_PulseWidth <= 174) Then SnapToMOPA = 150
	ElseIf (n_PulseWidth >= 175) And (n_PulseWidth <= 224) Then SnapToMOPA = 200
	ElseIf (n_PulseWidth >= 225) And (n_PulseWidth <= 274) Then SnapToMOPA = 250
	ElseIf (n_PulseWidth >= 275) And (n_PulseWidth <= 324) Then SnapToMOPA = 300
	ElseIf (n_PulseWidth >= 325) And (n_PulseWidth <= 374) Then SnapToMOPA = 350
	ElseIf (n_PulseWidth >= 375) And (n_PulseWidth <= 424) Then SnapToMOPA = 400
	ElseIf (n_PulseWidth >= 425) And (n_PulseWidth <= 474) Then SnapToMOPA = 450
	ElseIf (n_PulseWidth >= 475) And (n_PulseWidth <= 500) Then SnapToMOPA = 500
	Else
			SnapToMOPA = 0
	End If
End Function