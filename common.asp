<%
'  common.asp
'-------------------------------------------------------------------------------
'  Feature		: ASP Common Function Pack
'  Version		: v0.9
'  Author		: zhousong(zsol@qq.com)
'  Create Date	: 2008/2/11
'  Update Date	: 2012/4/28
'-------------------------------------------------------------------------------


'定义变量
dim Conn
dim Rs
dim SQL


'---------- DB操作相关函数------------------------------------------------------

'打开主数据库链接，ConnectionString可在外部配置文件中定义或本文件中定义
Sub OpenDB()
	On error Resume next
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open ConnectionString
	If Err.number <> 0 Then
		Response.Write "数据库服务器端连接错误，请检查数据库连接。"
		Response.Write Err.Description
		Err.Clear
		Conn.Close
		Set Conn = Nothing
		Set Rs = Nothing
	End If
End Sub


'关闭数据库链接
Sub CloseDB()
	Set Rs = Nothing
	If Conn.State = 1 Then Conn.Close()
	Set Conn = Nothing
End Sub


' 生成分页查询SQL语句
' 参数说明  
' ZD:字段列表 BM:表名 TJ:查询条件 PX:排序字段 Pagesize:每页记录数 PageNum:页码
Function BuildSQL(ZD,BM,TJ,PX,PageSize,PageNum)
	dim tmpValue
	IF CINT(PageNum) = 1 Then
		tmpValue = "SELECT TOP " & PageSize & " " & ZD & " FROM " & BM & " "
		IF TJ <> "" Then
			tmpValue = tmpValue & " WHERE " & TJ
		End IF
		tmpValue = tmpValue & " ORDER BY " & PX & " DESC"
	Else
		tmpValue = "SELECT " & ZD & " FROM " & BM & " WHERE " & PX & _
		" IN (SELECT TOP " & PageSize & " " & PX & " FROM (SELECT TOP " & _
		CSTR(PageSize * PageNum) & " " & PX & " FROM " & BM 
		IF TJ <> "" Then
			tmpValue = tmpValue & " WHERE " & TJ
		End IF
		tmpValue = tmpValue & " ORDER BY " & PX & " DESC) t1 ORDER BY " & _
		PX & " ASC) ORDER BY " & PX & " DESC"
	End IF
	BuildSQL = tmpValue
End Function


' 执行SQL,不返回记录集
Function ExecuteSQL(strSQL)
	OpenDB
	Conn.execute strSQL
	CloseDB
End Function


' 执行SQL,返回单个值
Function ExecuteScalar(strSQL)
	OpenDB
	Set Rs = Conn.execute(strSQL)
	IF Rs.BOF And Rs.EOF Then
		ExecuteScalar = Empty
	Else
		ExecuteScalar = Rs(0)
	End IF
	Rs.Close
	CloseDB
End Function


' 执行SQL,返回记录集数组
' 注：单列值查询的也返回二维数组，如a(0,0),a(0,1),a(0,2)...
Function ExecuteArray(strSQL)
	OpenDB
	Rs.Open strSQL,Conn,1,1
	ExecuteArray = Rs.GetRows
	Rs.Close
	CloseDB
End Function


' 执行SQL,返回记录集,用strFormat的内容格式化,模板中用{0},{1}...序列表示Rs的字段
Function ExecuteRs(strSQL,strFormat)
	dim i
	dim iFieldCount
	dim tmpValue
	dim tmpFormat
	tmpValue = ""
	tmpFormat = strFormat
	OpenDB
	Rs.Open strSQL,Conn,1,1
	IF Rs.EOF Then
		tmpValue = ""
	Else
		iFieldCount = Rs.Fields.Count
		Do Until Rs.EOF
			tmpFormat = strFormat

			' 下行用于将ID替换为链接地址
			'tmpFormat = Replace( tmpFormat,"{link}",LinkPath("detail",Rs(0),0) )
			
			For i = 0 to iFieldCount - 1
				tmpFormat = Replace(tmpFormat,"{" & CSTR(i) & "}",Rs(i))
			Next
			tmpValue = tmpValue & tmpFormat
			Rs.MoveNext
		Loop
	End IF
	Rs.Close
	CloseDB
	ExecuteRs = tmpValue
End Function




'---------- IO操作相关函数 -----------------------------------------------------

' 返回安全的SQL字符串
Function SafeSQL(strSQL)
	strSQL = Trim("" & strSQL)
	strSQL = Replace(Replace(Replace(strSQL,";","&#59;"),"'","''"),"-","&#45;")
	SafeSQL = strSQL
End Function


' 取参数值
Function GetRequest(RequestName)
	dim tmpValue
	tmpValue = "" & Request(RequestName)
	tmpValue = Server.HTMLEncode(tmpValue)
	tmpValue = SafeSQL(tmpValue)
	GetRequest = tmpValue
End Function


' 取数字型参数的值,如为空或不为数值则设为0
Function GetRequestNum(RequestName)
	dim tmpValue
	tmpValue = "" & Request(RequestName)
	IF tmpValue = "" OR NOT IsNumeric(tmpValue) Then
		tmpValue = 0
	Else
		tmpValue = Clng(tmpValue)
		if tmpValue < 0 then tmpValue = 0
	End IF
	GetRequestNum = tmpValue
End Function


'写文本文件
Function WriteFile(filename,text)
	dim txtFile,FSO
	IF(left(filename),1)="/" Then
		filename = server.Mappath(filename)
	End IF
	
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set txtFile = fso.CreateTextFile(filename)
	txtFile.Write text
	txtFile.Close
	Set txtFile = Nothing
	Set FSO = Nothing		
End Function


'读文本文件
Function ReadFile(filename)
	dim txtFile,FSO,tmp
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	IF(left(filename),1)="/" Then
		filename = server.Mappath(filename)
	End IF
	If  Fso.FileExists(filename) Then
		Set txtFile = FSO.OpenTextFile(filename,1,false)
		tmp = txtFile.ReadALL()
		txtFile.close
		Set txtFile = Nothing
	Else
		tmp = "文件未找到！"
	End IF
	Set FSO = Nothing
	ReadFile = tmp
End Function


Sub Js(ByVal Str)
	Response.Write("<sc" & "ript type=""text/javascript"">" & VbCrLf)
	Response.Write(VbTab & Str & VbCrLf)
	Response.Write("</scr" & "ipt>" & VbCrLf)
End Sub


Sub Alert(ByVal str)
	Response.Write("<sc" & "ript type=""text/javascript"">alert('" & JsEncode(str) & "\t\t');history.go(-1);</sc" & "ript>"&VbCrLf)
	Response.End()
End Sub


Sub AlertUrl(ByVal str, ByVal url)
	Response.Write("<sc" & "ript type=""text/javascript"">"&VbCrLf)
	Response.Write(VbTab&"alert('" & JsEncode(str) & "\t\t');location.href='" & url & "';"&VbCrLf)
	Response.Write("</sc" & "ript>"&VbCrLf)
	Response.End()
End Sub


Sub ConfirmUrl(ByVal str, ByVal Turl, ByVal Furl)
	Response.Write("<sc" & "ript type=""text/javascript"">"&VbCrLf)
	Response.Write(VbTab&"if(confirm('" & JsEncode(str) & "\t\t')){location.href='" & Turl & "';}else{location.href='" & Furl & "';}"&VbCrLf)
	Response.Write("</sc" & "ript>"&VbCrLf)
	Response.End()
End Sub


'函数名称:TextRead
'作用:利用AdoDb.Stream对象来读取UTF-8格式的文本文件
'参数:filename-文件物理路径;CharSet-编码格式(utf-8,gb2312.....)
Function TextRead(filename,CharSet)
	Dim str,stm
	Set stm = server.CreateObject("adodb.stream")
	stm.Type = 2 '文本模式读取
	stm.Mode = 3 
	stm.Charset = CharSet
	stm.Open
	stm.LoadFromFile filename
	str = stm.readtext
	stm.Close
	Set stm = Nothing
	TextRead = str
End Function


'函数名称:TextWrite
'作用:利用AdoDb.Stream对象来写入UTF-8格式的文本文件
'参数:filename-文件物理路径;Str-文件内容;CharSet-编码格式(utf-8,gb2312.....)
Function TextWrite(filename,byval Str,CharSet) 
	Dim stm
	Set stm = Server.CreateObject("adodb.stream")
	stm.Type = 2 '文本模式
	stm.mode = 3
	stm.Charset = CharSet
	stm.open
	stm.WriteText str
	stm.SaveToFile filename,2 
	stm.Flush
	stm.Close
	set stm = Nothing
End Function




'---------- 调试用函数 -------------------------------------------------

dim TestTime1	' 测试程序运行时间用,程序开始运行时间
dim TestTime2	' 测试程序运行时间用,程序运行结束时间

' 默认自动初始化testTime1,只需在页尾调用t2即可。
' 需更精确地测试时,可以再调用t1,运行任务,t2
TestTime1 = timer()

' 测试程序运行时间
Sub t1()
	TestTime1 = timer()	
End Sub


Sub t2()
	TestTime2 = timer()
	Response.Write "<br>运行时间：" & FormatNumber(( TestTime2 - TestTime1 )*1000,3) & "ms<br>" 
End Sub


'调试变量
Function d(vName)
	Response.Write vName
	Response.Write "<br />"
	Response.flush()
End Function


' 列印出表单提交的参数值
Sub PR()
	dim a
	For Each a In Request.Form
		Response.write a
		Response.write ":"
		Response.write Request.Form(a)
		Response.write "<br>"
	Next
End Sub


' 列印出URL查询的参数值
Sub PQ()
	dim a
	For Each a In Request.QueryString
		Response.write a
		Response.write ":"
		Response.write Request.QueryString(a)
		Response.write "<br>"
	Next
End Sub


' 列印出Application变量
Sub PA()
	Dim a
	For Each a In Application.Contents
		Response.write a
		Response.write ":"
		Response.write Application.Contents(a)
		Response.write "<br>"
	Next
End Sub


' 列印出Session变量
Sub PS()
	dim a
	For Each a In Session.Contents
		Response.write a
		Response.write ":"
		Response.write Session.Contents(a)
		Response.write "<br>"
	Next
End Sub




'----------编码/解码函数--------------------------------------------------------
'HTML格式化
Function HtmlFormat(ByVal str)
	If Not IsN(str) Then
		Dim m : Set m = RegMatch(str, "<([^>]+)>")
		For Each Match In m
			 str = Replace(str, Match.SubMatches(0), regReplace(Match.SubMatches(0), "\s+", Chr(0)))
		Next
		Set m = Nothing
		str = Replace(str, Chr(32), "&nbsp;")
		str = Replace(str, Chr(9), "&nbsp;&nbsp; &nbsp;")
		str = Replace(str, Chr(0), " ")
		str = regReplace(str, "(<[^>]+>)\s+", "$1")
		str = Replace(str, vbCrLf, "<br />")
	End If
	HtmlFormat = str
End Function


'HTML编码
Function HtmlEncode(ByVal str)
	If Not IsN(str) Then
		str = Replace(str, Chr(38), "&#38;")
		str = Replace(str, "<", "&lt;")
		str = Replace(str, ">", "&gt;")
		str = Replace(str, Chr(39), "&#39;")
		str = Replace(str, Chr(32), "&nbsp;")
		str = Replace(str, Chr(34), "&quot;")
		str = Replace(str, Chr(9), "&nbsp;&nbsp; &nbsp;")
		str = Replace(str, vbCrLf, "<br />")
	End If
	HtmlEncode = str
End Function


'HTML解码
Function HtmlDecode(ByVal str)
	If Not IsN(str) Then
		str = regReplace(str, "<br\s*/?\s*>", vbCrLf)
		str = Replace(str, "&nbsp;&nbsp; &nbsp;", Chr(9))
		str = Replace(str, "&quot;", Chr(34))
		str = Replace(str, "&nbsp;", Chr(32))
		str = Replace(str, "&#39;", Chr(39))
		str = Replace(str, "&apos;", Chr(39))
		str = Replace(str, "&gt;", ">")
		str = Replace(str, "&lt;", "<")
		str = Replace(str, "&amp;", Chr(38))
		str = Replace(str, "&#38;", Chr(38))
		HtmlDecode = str
	End If
End Function


' HTML过滤
Function HtmlFilter(ByVal str)
	str = regReplace(str,"<[^>]+>","")
	str = Replace(str, ">", "&gt;")
	str = Replace(str, "<", "&lt;")
	HtmlFilter = str
End Function


' JS编码
Function JsEncode(ByVal str)
	If Not isN(str) Then
		str = Replace(str,Chr(92),"\\")
		str = Replace(str,Chr(34),"\""")
		str = Replace(str,Chr(39),"\'")
		str = Replace(str,Chr(9),"\t")
		str = Replace(str,Chr(13),"\r")
		str = Replace(str,Chr(10),"\n")
		str = Replace(str,Chr(12),"\f")
		str = Replace(str,Chr(8),"\b")
	End If
	JsEncode = str
End Function


' ESCAPE
Function Escape(ByVal str)
	Dim i,c,a,s : s = ""
	If isN(str) Then Escape = "" : Exit Function
	For i = 1 To Len(str)
		c = Mid(str,i,1)
		a = ASCW(c)
		If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
			s = s & c
		ElseIf InStr("@*_+-./",c)>0 Then
			s = s & c
		ElseIf a>0 and a<16 Then
			s = s & "%0" & Hex(a)
		ElseIf a>=16 and a<256 Then
			s = s & "%" & Hex(a)
		Else
			s = s & "%u" & Hex(a)
		End If
	Next
	Escape = s
End Function


' UNESCAPE
Function UnEscape(ByVal str)
	Dim x, s
	x = InStr(str,"%")
	s = ""
	Do While x>0
		s = s & Mid(str,1,x-1)
		If LCase(Mid(str,x+1,1))="u" Then
			s = s & ChrW(CLng("&H"&Mid(str,x+2,4)))
			str = Mid(str,x+6)
		Else
			s = s & Chr(CLng("&H"&Mid(str,x+1,2)))
			str = Mid(str,x+3)
		End If
		x=InStr(str,"%")
	Loop
	UnEscape = s & str
End Function




'----------其它工具函数---------------------------------------------------------

Function IIF(expr, truepart, falsepart)
	IF expr = False Then
		IIF = falsepart
	Else
		IIF = truepart
	End IF
End Function


Function isN(ByVal str)
	isN = False
	Select Case VarType(str)
		Case vbEmpty, vbNull
			isN = True
			Exit Function
		Case vbString
			If str="" Then isN = True
			Exit Function
		Case vbObject
			If TypeName(str)="Nothing" Or TypeName(str)="Empty" Then
				isN = True
			End IF
			Exit Function
		Case vbArray,8194,8204,8209
			If Ubound(str)=-1 Then isN = True
			Exit Function
	End Select
End Function


' 日期格式化函数
' 参数 strdate:要格式化的日期，fstr:格式字符串
Function DateFormat(strDate,fstr)
	IF isdate(strDate) Then
		Dim i,temp
		temp=replace(fstr,"yyyy",DatePart("yyyy",strDate))
		temp=replace(temp,"yy",mid(DatePart("yyyy",strDate),3))
		temp=replace(temp,"y",DatePart("y",strDate))
		temp=replace(temp,"w",DatePart("w",strDate))
		temp=replace(temp,"ww",DatePart("ww",strDate))
		temp=replace(temp,"q",DatePart("q",strDate))
		temp=replace(temp,"mm",iif(len(DatePart("m",strDate))>1,DatePart("m",strDate),"0"&DatePart("m",strDate)))
		temp=replace(temp,"dd",iif(len(DatePart("d",strDate))>1,DatePart("d",strDate),"0"&DatePart("d",strDate)))
		temp=replace(temp,"hh",iif(len(DatePart("h",strDate))>1,DatePart("h",strDate),"0"&DatePart("h",strDate)))
		temp=replace(temp,"nn",iif(len(DatePart("n",strDate))>1,DatePart("n",strDate),"0"&DatePart("n",strDate)))
		temp=replace(temp,"ss",iif(len(DatePart("s",strDate))>1,DatePart("s",strDate),"0"&DatePart("s",strDate)))
		DateFormat=temp
	Else
		DateFormat=false
	End IF
End Function


' 禁用缓存
Sub noCache()
	Response.Buffer = True
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
	Response.AddHeader "Expires",Date()
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
End Sub

' 正则检测
Function RegCheck(str,reg)
	Dim re
	Set re = New RegExp
	re.Pattern = reg
	re.Global = True
	re.IgnoreCase = True
	re.MultiLine = True
	RegCheck = re.Test(str)
End Function


' 正则替换
Function RegReplace(str,regFind,regRep)
	Dim re
	Set re = New RegExp
	re.Pattern = regFind
	re.Global = True
	re.IgnoreCase = True
	re.MultiLine = True
	RegReplace = re.Replace(str,regRep)
End Function


' 正则匹配
Function RegMatch(ByVal str, ByVal rule)
	Dim Reg
	Set Reg = New Regexp
	Reg.Global = True
	Reg.IgnoreCase = True
	Reg.Pattern = rule
	Set RegMatch = Reg.Execute(str)
	Set Reg = Nothing
End Function


' 常用格式正则检测函数
Function Test(ByVal Str, ByVal Pattern)
	Dim Pa
	Select Case Lcase(Pattern)
		Case "date"		Test = IIF(isDate(Str),True,False) : Exit Function
		Case "idcard"	Pa = "^\d{15}$)|(\d{17}(?:\d|x|X)$"
		Case "english"	Pa = "^[A-Za-z]+$"
		Case "chinese"	Pa = "^[\u0391-\uFFE5]+$"
		Case "username"	Pa = "^[a-z]\w{2,19}$"
		Case "email"	Pa = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
		Case "int"		Pa = "^[-\+]?\d+$"
		Case "number"	Pa = "^\d+$"
		Case "double"	Pa = "^[-\+]?\d+(\.\d+)?$"
		Case "price"	Pa = "^\d+(\.\d+)?$"
		Case "zip"		Pa = "^[1-9]\d{5}$"
		Case "qq"		Pa = "^[1-9]\d{4,9}$"
		Case "phone"	Pa = "^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$"
		Case "mobile"	Pa = "^((\(\d{2,3}\))|(\d{3}\-))?(1[35][0-9]|189)\d{8}$"
		Case "url"		Pa = "^(http|https|ftp):\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\""])*$"
		Case "domain"	Pa = "^[A-Za-z0-9\-]+\.([A-Za-z]{2,4}|[A-Za-z]{2,4}\.[A-Za-z]{2})$"
		Case "ip"		Pa = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
		Case Else Pa = Pattern
	End Select
	Test = RegCheck(CStr(Str),Pa)
End Function


' 检测提交页面来源，本机提交返回真，否则为假
Function CheckDataFrom()
	Dim v1, v2
	CheckDataFrom = False
	v1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
	v2 = Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(v1,8,Len(v2)) = v2 Then
		CheckDataFrom = True
	End If
end Function


' 取来访IP
Function GetIP()
	Dim addr, x, y
	x = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	y = Request.ServerVariables("REMOTE_ADDR")
	addr = IIF(isN(x) or lCase(x)="unknown",y,x)
	If InStr(addr,".")=0 Then addr = "0.0.0.0"
	GetIP = addr
End Function


'分页导航生成函数 参数：URL（有其它查询参数时,URL以&结尾，无其它查询参数时以？结尾），当前页数，总记录数
Function ListPage(PageURL,CurPage,PageCount)
	dim page_info
	page_info = "<div class=""pagelist"">"
	IF CurPage<1 Then CurPage=1
	IF CurPage>PageCount Then CurPage=PageCount
	If PageCount<=10 Then
		For i = 1 To PageCount
			IF CurPage = i Then
				page_info = page_info &  ("<span>" & i & "</span>")
			Else
				page_info = page_info &  ("<a href=""" & PageURL & "page=" & i & """>[" & i & "]</a>")
			End IF		
		Next
	Else
		If CurPage>6 Then
			page_info = page_info &  ("<a href=""" & PageURL & "page=1"">[1]</a>...")	
		End If
		
		If CurPage<6 Then
			StartPage = 1
			EndPage = 10
		Else
			StartPage = CurPage-5
		End If
		
		If CurPage+4>PageCount Then
			EndPage = PageCount
			StartPage = PageCount-10
		Else
			If CurPage>=6 Then
				EndPage = CurPage+4
			End IF
		End If

		For i = StartPage To EndPage
			IF (i = int(CurPage)) Then
				page_info = page_info &  ("<span>" & i & "</span>")
			Else
				page_info = page_info &  ("<a href=""" & PageURL & "page=" & i & """>[" & i & "]</a>")
			End IF
		Next

		If CurPage+4<PageCount Then
			page_info = page_info &  ("...<a href=""" & PageURL & "page=" & PageCount & """>[" & PageCount & "]</a>")
		End IF
	End IF
	page_info = page_info &  "</div>"	
	ListPage = page_info
End Function



%>
