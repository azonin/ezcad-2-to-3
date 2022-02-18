WScript.Echo Chr(16) & String(115,Chr(150)) & Chr(17)
WScript.Echo Chr(166) & " Convert EZCAD2 to EZCAD3 Marking Parameters                                                                       " & Chr(166)
WScript.Echo Chr(166) & " Copyright " & Chr(1) & " 2022 Alex Zonin                                                                                       " & Chr(166)
WScript.Echo Chr(166) & String(115,Chr(150)) & Chr(166)
WScript.Echo Chr(166) & " - Imports MarkParam.lib file                                                                                      " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - Converts the old EZCAD2 parameter format to the new EZCAD3                                                      " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - Saves the new format as MarkParamlib.ini (appends to the existing file if you put it in the same folder)        " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - All files have to be in the same folder as this script                                                          " & Chr(166)
WScript.Echo Chr(166) & String(115," ") & Chr(166)
WScript.Echo Chr(166) & " - If you want to change your time correction settings (TC values), you can do it in the section-template.txt file " & Chr(166)
WScript.Echo Chr(16) & String(115,Chr(150)) & Chr(17)
WScript.Echo ""

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
s_StartDir = fso.GetAbsolutePathName(".") & "\"

s_OldINIFile = s_StartDir + "MarkParam.lib"
s_NewINIFile = s_StartDir + "MarkParamlib.ini"
s_SectionTemplateFile = s_StartDir + "section-template.txt"

If fso.FileExists(s_SectionTemplateFile) Then
	Set obj_File = fso.OpenTextFile(s_SectionTemplateFile,1)
			s_SectionTemplate = obj_File.ReadAll
	Set obj_File = Nothing
End If

If fso.FileExists(s_OldINIFile) Then
	Set obj_File = fso.OpenTextFile(s_OldINIFile,1)
		Do While obj_File.AtEndOfStream <> True
			s_TempBuff = obj_File.ReadLine

			If s_TempBuff = "[LASERMODE]" Then
				Do While Len(s_TempBuff) > 1
					s_TempBuff = obj_File.ReadLine
				Loop

			ElseIf Left(s_TempBuff,1)="[" Then
				'Initialize the Target Section
				s_TargetSection = s_SectionTemplate
				
				'Extract the Section Name
				s_SectionName = Mid(s_TempBuff,2,Len(s_TempBuff)-2)
				
				s_TargetSection = Replace(s_TargetSection,"[SECTION-NAME]","[" & s_SectionName & "]")
				s_TargetSection = Replace(s_TargetSection,"NAME=NAME","NAME=" & s_SectionName)
				s_TargetSection = Replace(s_TargetSection,"DESC=DESC","DESC=Put your description right here")

			ElseIf Left(s_TempBuff,10) = "MARKSPEED=" Then
				s_TargetSection = Replace(s_TargetSection,"MARKSPEED=MARKSPEED","MARKSPEED=" & FormatNumber(CDbl(Right(s_TempBuff,Len(s_TempBuff)-10)),6,-1,0,0))

			ElseIf Left(s_TempBuff,11) = "POWERRATIO=" Then
				s_TargetSection = Replace(s_TargetSection,"POWERRATIO=POWERRATIO","POWERRATIO=" & FormatNumber(CDbl(Right(s_TempBuff,Len(s_TempBuff)-11)),6,-1,0,0))

			ElseIf Left(s_TempBuff,12) = "QPULSEWIDTH=" Then
				s_TargetSection = Replace(s_TargetSection,"QPULSEWIDTH=QPULSEWIDTH","QPULSEWIDTH=" & FormatNumber(CDbl(Right(s_TempBuff,Len(s_TempBuff)-12)),6,-1,0,0))
			
			ElseIf Left(s_TempBuff,11) = "WOBBLEMODE=" Then
				s_TargetSection = Replace(s_TargetSection,"WOBBLEMODE=WOBBLEMODE","WOBBLEMODE=" & Right(s_TempBuff,Len(s_TempBuff)-11))
			
			ElseIf Left(s_TempBuff,15) = "WOBBLEDIAMETER=" Then
				s_TargetSection = Replace(s_TargetSection,"WOBBLEDIAMETER=WOBBLEDIAMETER","WOBBLEDIAMETER=" & FormatNumber(CDbl(Right(s_TempBuff,Len(s_TempBuff)-15)),6,-1,0,0))
		
			ElseIf Left(s_TempBuff,11) = "WOBBLEDIST=" Then
				s_TargetSection = Replace(s_TargetSection,"WOBBLEDIST=WOBBLEDIST","WOBBLEDIST=" & FormatNumber(CDbl(Right(s_TempBuff,Len(s_TempBuff)-11)),6,-1,0,0))
	
			ElseIf Left(s_TempBuff,5) = "FREQ=" Then
				s_TargetSection = Replace(s_TargetSection,"FREQF=FREQF","FREQF=" & FormatNumber(CDbl(Right(s_TempBuff,Len(s_TempBuff)-5)),6,-1,0,0))

			ElseIf Len(s_TempBuff)<=1 Then
				Set obj_OutputFile = fso.OpenTextFile(s_NewINIFile, 8, True) 'Appends to your existing MarkParamlib.ini if there is one in the same folder as this script
				obj_OutputFile.Write s_TargetSection
				obj_OutputFile.Write vbCrLf
				Set obj_OutputFile = Nothing				

			End If

		Loop
	Set obj_File = Nothing
End If

Set fso = Nothing