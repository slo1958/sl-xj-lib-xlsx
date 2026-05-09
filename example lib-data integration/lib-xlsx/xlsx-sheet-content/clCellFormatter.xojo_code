#tag Class
Protected Class clCellFormatter
	#tag Method, Flags = &h1
		Protected Sub checkDateFormat(format as String)
		  var testStrs() as string
		  
		  testStrs = array("y","yy","yyyy", "m", "mm", "mmm", "d", "dd", "ddd")
		  
		  var cnt as integer
		  
		  for each teststr as string in testStrs
		    if format.IndexOf(teststr) >= 0 then cnt = cnt +1
		    
		  next
		  
		  self.IsDateFormat = cnt > 1
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(format as string)
		  
		  self.SourceFormat = format
		  
		  self.checkDateFormat(format)
		  
		  if self.IsDateFormat then
		    self.PrepareDateFormat(format)
		    
		  else
		    self.PrepareNumberFormat(format)
		    
		  end if
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function DayName(DayOfWeek as integer) As string
		  
		  Select Case DayOfWeek
		  Case 1
		    Return "Sunday"
		  Case 2
		    Return "Monday"
		  Case 3
		    Return "Tuesday"
		  Case 4
		    Return "Wednesday"
		  Case 5
		    Return "Thursday"
		  Case 6
		    Return "Friday"
		  Case 7
		    Return "Saturday"
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FormatAsDate(d as DateTime) As string
		  
		  var ret as string
		  
		  
		  if sections.LastIndex < 0 then return d.SQLDate
		  
		  if sections.LastIndex >= 0 and sections(0) = "@" then return d.SQLDate
		  
		  for each s as string in self.Sections
		    select case s
		    case "yy"
		      ret = ret + format(d.Year, "####").right(2)
		      
		    case "yyyy"
		      ret = ret + format(d.Year, "####")
		      
		    case "m"
		      ret = ret + str(d.month)
		    case "mm"
		      ret = ret + format(d.month, "00")
		      
		    case "mmm"
		      ret = ret + MonthName(d.month).Uppercase.left(3)
		      
		    case "mmmm"
		      ret = ret + MonthName(d.month).Titlecase()
		      
		    case "d"
		      ret = ret + str(d.day)
		      
		    case "dd"
		      ret = ret + format(d.day, "00")
		      
		    case "dddd"
		      ret = ret + DayName(d.DayOfWeek)
		      
		    case "e"
		      ret = ret + "?e"
		      
		    case else
		      ret = ret + s
		      
		    end Select
		    
		  next
		  
		  Return ret
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FormatAsNumber(d as double) As string
		  
		  const CDefaultNumberFormat = "-#####0.00##"
		  
		  var format as string
		  
		  if Sections.LastIndex >= 0 and sections(0).Length>1 then
		    format = sections(0)
		    
		  else
		    format = "General"
		    
		  end if
		  
		  
		  if format = "General" or format = "@" then
		    return format(d, CDefaultNumberFormat)
		    
		  else
		    return format(d, format)
		    
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FormatString() As string
		  Return SourceFormat
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FormatValue(SourceValue as variant) As string
		  
		  if self.IsDateFormat then
		    
		    var d as DateTime = MakeDate(SourceValue)
		    // 
		    // var dateOffset as integer = SourceValue
		    // var d as new DateTime(new Date(1900,1,1))
		    // 
		    // d = d.AddInterval(0,0, dateOffset-2)
		    // 
		    return self.FormatAsDate(d)
		  else
		    var d as Double = SourceValue
		    return self.FormatAsNumber(d)
		    
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MakeDate(SourceValue as variant) As DateTime
		  
		  
		  var dateOffset as integer = SourceValue
		  var d as new DateTime(new Date(1900,1,1))
		  
		  d = d.AddInterval(0,0, dateOffset-2)
		  
		  return d
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function MonthName(Month as integer) As string
		  Select Case Month
		    
		  Case 1
		    Return "January"
		    
		  Case 2
		    Return "February"
		    
		  Case 3
		    Return "March"
		    
		  Case 4
		    Return "April"
		    
		  Case 5
		    Return "May"
		    
		  Case 6
		    Return "June"
		    
		  Case 7
		    Return "July"
		    
		  Case 8
		    Return "August"
		    
		  Case 9
		    Return "September"
		    
		  Case 10
		    Return "October"
		    
		  Case 11
		    Return "November"
		    
		  Case 12
		    Return "December"
		    
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub PrepareDateFormat(format as string)
		  
		  
		  var f2 as string = format
		  
		  
		  if f2.IndexOf(";") >=0  then
		    var f3() as string = f2.Split(";")
		    
		    f2 = ""
		    
		    for each f3element as string in f3
		      if f3element.left(1)="[" then
		        
		      elseif f2 = "" then
		        f2 = f3element
		        
		      end if
		    next
		    
		    if f2.Length = 0 then f2 = format
		    
		  end if
		  
		  
		  if f2.IndexOf("[$") >= 0 then
		    var p1 as integer = f2.IndexOf("[")
		    var p2 as integer = f2.IndexOf("]")
		    
		    var slcid as String = f2.Middle(p1+3, p2-p1-3)
		    var sformat as string = f2.Middle(p2+1,999)
		    
		    f2 = sformat
		    
		  end if
		  
		  for i as integer = 0 to f2.Length
		    var c as String = f2.Middle(i,1)
		    
		    if sections.LastIndex < 0 then
		      sections.add(c)
		      
		    elseif sections(sections.LastIndex).left(1) = c then
		      sections(sections.LastIndex) = sections(sections.LastIndex) + c
		      
		    elseif c="\" then 
		      
		    else
		      sections.add(c)
		      
		    end if
		    
		  next
		  
		  return 
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub PrepareNumberFormat(format as string)
		  
		  var f2 as string = format
		  
		  Sections.RemoveAll
		  
		  if f2.IndexOf(";") >=0  then
		    var f3() as string = f2.Split(";")
		    
		    for each f3element as string in f3
		      if f3element.left(2)="[$" then // remove local mark
		        var idx as integer = f3element.IndexOf("]")
		        
		        Sections.Add( f3element.mid(idx+1, 999))
		        
		      else
		        Sections.Add( f3element.trim)
		        
		      end if
		      
		    next
		    return
		    
		  end if
		  
		  sections.Add(format)
		  
		  return
		  
		  
		End Sub
	#tag EndMethod


	#tag Note, Name = Decsription
		
		Format type    Meaning
		d    Day of the month as digits without leading zeros for single-digit days.
		dd    Day of the month as digits with leading zeros for single-digit days.
		ddd    Abbreviated day of the week as specified by a LOCALE_SABBREVDAYNAME* value, for example, "Mon" in English (United States).Windows Vista and later: If a short version of the day of the week is required, your application should use the LOCALE_SSHORTESTDAYNAME* constants.
		dddd    Day of the week as specified by a LOCALE_SDAYNAME* value.
		
		
		Format type    Meaning
		M    Month as digits without leading zeros for single-digit months.
		MM    Month as digits with leading zeros for single-digit months.
		MMM    Abbreviated month as specified by a LOCALE_SABBREVMONTHNAME* value, for example, "Nov" in English (United States).
		MMMM    Month as specified by a LOCALE_SMONTHNAME* value, for example, "November" for English (United States), and "Noviembre" for Spanish (Spain).
		
		
		Format type    Meaning
		y    Year represented only by the last digit.
		yy    Year represented only by the last two digits. A leading zero is added for single-digit years.
		yyyy    Year represented by a full four or five digits, depending on the calendar used. Thai Buddhist and Korean calendars have five-digit years. The "yyyy" pattern shows five digits for these two calendars, and four digits for all other supported calendars. Calendars that have single-digit or two-digit years, such as for the Japanese Emperor era, are represented differently. A single-digit year is represented with a leading zero, for example, "03". A two-digit year is represented with two digits, for example, "13". No additional leading zeros are displayed.
		yyyyy    Behaves identically to "yyyy".
		
		
		Format type    Meaning
		g, gg    Period/era string formatted as specified by the CAL_SERASTRING value. The "g" and "gg" format pictures in a date string are ignored if there is no associated era or period string.
		
		
	#tag EndNote

	#tag Note, Name = Documentation
		
		
		Source: https://stackoverflow.com/questions/54134729/what-does-the-130000-in-excel-locale-code-130000-mean/54540455#54540455
		
		Referenced:
		https://learn.microsoft.com/en-us/windows/win32/intl/language-identifier-constants-and-strings
		https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/63d3d639-7fd2-4afb-abbe-0d5b5551eef8?redirectedfrom=MSDN
		https://learn.microsoft.com/en-us/windows/win32/intl/calendar-identifiers
		https://learn.microsoft.com/en-us/windows/win32/intl/national-language-support-constants
		
		
		
		The correct format is technically... xxyyzzzz
		
		=TEXT(A1,"[$-130000]d/m/yyyy")
		
		Where does $-130000 come from? Is this an Excel specific locale function?
		xx = 00 if missing
		
		
		xx = 00
		
		yy = 13
		
		zzzz = 0000
		
		xx: the first two digits (xx) represent the appearance of the number
		Hexadecimal value = Reserved Bit for Application Use (Application Specific - From what I have been reading)
		00 = System Defaults (Set in Control panel)
		01 = Western language
		02 = Arabic Hindi
		03 = Extend Arabic Hindi
		04 = Sanskrit
		05 = Bengali
		06 = Gorumuchi
		07 = Gujarati
		08 = Oriya
		09 = Tamil
		0A = Telugu
		0B = Kannada
		0C = Malayalam
		0D = Thai
		0E = Laotian
		0F = Tibetan language
		10 = Burmese
		11 = Ethiopian
		12 = Cambodian
		13 = Mongolian
		1B = Japanese 1
		1C = Japanese 2
		1D = Japanese 3
		1E = Simplified Chinese 1, Chinese lowercase
		1F = Simplified Chinese 2, Chinese uppercase
		20 = Simplified Chinese 3, full-width numbers
		21 = Traditional Chinese 1, traditional lowercase
		22 = Traditional Chinese 2, traditional uppercase
		23 = Traditional Chinese 3, full-width numbers
		24 = Korean 1
		25 = Korean 2
		26 = Korean 3
		27 = Korean 4
		
		
		yy: represents the calendar format (xxyyzzzz):
		Hexadecimal value = Calendar ID / Sort ID = See Library Source Below
		00 = System Defaults (Set in Control panel)
		01 = Gregorian calendar (localization)
		02 = Gregorian calendar (United States)
		03 = Japanese calendar (and calendar)
		04 = Taiwan calendar
		05 = Korean calendar (Tanji)
		06 = Hajj (Arab Lunar Calendar)
		07 = Thai
		08 = Jewish Calendar
		09 = Gregorian calendar (Middle Eastern French)
		11 = Lunar Calendar (Not Officially published)
		12 = Lunar Calendar (Not Officially published)
		13 = Lunar Calendar (Not Officially published)
		0A = Gregorian calendar (Arabic)
		0B = Gregorian calendar (translated English)
		0E = Lunar Calendar (Not Officially published)
		
		
		zzzz:represent the language code (xxyyzzzz):
		Hexadecimal value = Language ID Values= LCID
		0000 = System Defaults (Set in Control panel) = Not certain if the "control panel" has an LCID
		0401 = Arabic = 1025
		0402 = Bulgarian = 1026
		0403 = Catalan = 1027
		0404 = traditional Chinese) = 1028
		0405 = Czech = 1029
		0406 = Danish = 1030
		0407 = German = 1031
		0408 = Greek = 1032
		0409 = English (United States) = 1033
		040B = Finnish = 1035
		040C = French = 1036
		040D = Hebrew = 1037
		040E = Hungarian = 1038
		040F = Icelandic = 1039
		0410 = Italian = 1040
		0411 = Japanese = 1041
		0412 = Korean = 1042
		0413 = Dutch = 1043
		0414 = Norwegian (Birkmer) = 1044
		0415 = Polish = 1045
		0416 = Portuguese (Brazil) = 1046
		0418 = Romanian = 1048
		0419 = Russian = 1049
		041A = Croatian = 1050
		041B = Slovak = 1051
		041C = Albanian = 1052
		041D = Swedish = 1053
		041E = Thai = 1054
		041F = Turkish = 1055
		0420 = Urdu = 1056
		0421 = Indonesian = 1057
		0422 = Ukrainian = 1058
		0423 = Belarusian = 1059
		0424 = Slovenian = 1060
		0425 = Estonian = 1061
		0426 = Latvian = 1062
		0427 = Lithuanian = 1063
		0428 = Tajik = 1064
		0429 = Persian = 1065
		042A = Vietnamese = 1066
		042B = Armenian = 1067
		042C = Azerbaijani (Latin) = 1068
		042D = Basque = 1069
		042F = Macedonian = 1071
		0436 = Afrikaans = 1078
		0437 = Georgian = 1079
		0438 = Faroese = 1080
		0439 = Hindi = 1081
		043A = Maltese = 1082
		043D = Yiddish = 1085
		043E = Malay = 1086
		043F = Kazakh = 1087
		0440 = Kyrgyz = 1088
		0441 = Swahili = 1089
		0442 = Turkmen = 1090
		0443 = Uzbek (Latin) = 1091
		0444 = Proverb = 1092
		0445 = Bengali = 1093
		0446 = Punjabi = 1094
		0447 = Gujarati = 1095
		0448 = Oriya = 1096
		0449 = Tamil = 1097
		044A = Telugu = 1098
		044B = Kannada = 1099
		044C = Malayalam = 1100
		044D = Assamese = 1101
		044E = Marathi = 1102
		044F = Sanskrit = 1103
		0450 = Mongolian = 1104
		0456 = Galician = 1110
		0457 = Gungan = 1111
		0458 = Manipur = 1112
		0459 = Sindhi = 1113
		045A = Syrian = 1114
		045B = Sinhalese = 1115
		045C = Congga = 1116
		045D = Inuit = 1117
		045E = Amharic = 1118
		045F = Tamasic (Berber / Arab) = 1119
		0460 = Kashmiri (Arabic) = 1120
		0461 = Nepali = 1121
		0462 = Frisian = 1122
		0463 = Pashto = 1123
		0464 = Filipino = 1124
		0465 = Dhivehi = 1125
		0466 = Kwa = 1126
		0467 = Furbe = 1127
		0468 = Hausa = 1128
		0469 = Ibibio = 1129
		046A = Yoruba = 1130
		0470 = Igbo = 1136
		0471 = Kanuri = 1137
		0472 = Kucht = 1138
		0473 = Tigrinya (Ethiopia) = 1139
		0475 = Hawaiian = 1141
		0476 = Latin = 1142
		0477 = Somali = 1143
		0478 = Proverb = 1144
		0804 = Chinese (Simplified) = 2052
		0807 = German (Switzerland) = 2055
		0809 = English (UK) = 2057
		0814 = Norwegian (Nynorsk) = 2068
		0816 = Portuguese (Portugal) = 2070
		081A = Serbian (Latin) = 2074
		082C = Azeri (Cyrillic) = 2092
		0843 = Uzbek (Cyrillic) = 2115
		0873 = Tigrinya (Eritrea) = 2163
		085F = Tamasic (Latin) = 2143
		0C07 = German (Austria) = 3079
		0C09 = English (Australia) = 3081
		0C0A = Spanish = 3082
		0C0C = French (Canada) = 3084
		0C1A = Serbian (Cyrillic) = 3098
		1009 = English (Canada) = 4105
		
		
		
		
		
	#tag EndNote

	#tag Note, Name = Examples
		
		Note that in 2023, Excel 365 shows automatically [$-en-US] instead of [$-409] ([$-zh-CN] instead of [$-804], etc.)
		
		[$-409]mmmm dd yyyy  h:mm AM/PM
		November 27 1973  11:28 AM
		
		
		[$-804]mmmm dd yyyy  h:mm AM/PM
		十一月 27 1973  11:28 上午
		
		Format Code  409 (English United States)  804 (Chinese PRC)
		===========  ===========================  =================
		m            11                           11
		mm           11                           11
		mmm          Nov                          十一月
		mmmm         November                     十一月
		d            27                           27
		dd           27                           27
		ddd          Mon                          二
		dddd         Monday                       星期二
		y            73                           73
		yy           73                           73
		yyy          1973                         1973
		yyyy         1973                         1973
		AM/PM        AM                           上午
		
		
	#tag EndNote


	#tag Property, Flags = &h0
		IsDateFormat As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Sections() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		SourceFormat As String
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="SourceFormat"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="IsDateFormat"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
