<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
Dim rs
Dim sql
Dim l_id
Dim l_tipo
Dim oScript
Dim strCMD
Dim RetCode

l_id  = request.QueryString("id")
l_tipo  = request.QueryString("tipo")

on error goto 0
Set rs = Server.CreateObject("ADODB.RecordSet")
sql = "SELECT servdesabr "
sql = sql & "FROM rhproappsrv where id=" & l_id
rsOpen rs, cn, sql, 0 

if not rs.eof then
	if l_tipo = "S" then
		'strCMD = "%comspec% cmd /c  sol /c" 
		'strCMD = "cmd /c NET START " & chr(34) & trim(rs(0)) & chr(34)
		strCMD = "cmd /c NET START " & chr(34) & "RHPro App Server" & chr(34)
		'strCMD = "cmd /C NET START tlntsvr"
		'strCMD = "c:\ticket\cgi-bin\servicios\startService.cmd"
	else
		strCMD = " cmd /c NET STOP " & chr(34) & trim(rs(0)) & chr(34)
	end if
	'on error resume next
	Set oScript = Server.CreateObject("WScript.Shell")
	Response.Write strCMD & vbcrlf
	RetCode = oScript.run (strCMD,0,true)'<----------
	Response.Write err.description
	Response.Write "ret==" & RetCode

	'set RetCode = oScript.exec(strCMD)'<----------
	'Do While RetCode.Status = 0
	'	Response.Write RetCode.status & vbcrlf
'	'   'wshell.Sleep 100
	'Loop
	'Response.write "<script>alert('return="& RetCode.StdOut.ReadAll &"');</script>"
	'set wshell = nothing
	Response.Write "-err.num=" & err.number
	if RetCode = 0 Then'.status
		if l_tipo = "S" then
			Response.write "<script>alert('Proceso Iniciado.');//window.close();</script>"
		else
			Response.write "<script>alert('Proceso Detenido.');//window.close();</script>"
		end if
	else
		Response.write "<script>alert('Error al querer iniciar/detener.');//window.close();</script>"
	end if
end if
rs.Close
set rs = Nothing
set oScript = Nothing
cn.Close
set cn = Nothing
%>
