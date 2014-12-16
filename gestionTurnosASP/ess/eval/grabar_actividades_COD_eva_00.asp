<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_aspectomejorar
  Dim l_planaccion
  Dim l_planfecharev
  Dim l_evaplnro
  
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
  Dim l_tipo  
' parametros de entrada
  l_aspectomejorar = left(trim(request.querystring("aspectomejorar")),200)
  l_planaccion     = left(trim(request.querystring("planaccion")),200)
  l_planfecharev   = request.querystring("planfecharev")
  l_evldrnro = request.querystring("evldrnro")
  l_tipo = request.querystring("tipo")
  l_evaplnro = request.querystring("evaplnro")
  

'BODY ----------------------------------------------------------
if l_tipo ="A" then
	l_sql= "insert into evaplan (aspectomejorar,planaccion,planfecharev,evldrnro) "
	l_sql = l_sql & " values ('" & trim(l_aspectomejorar) & "','" & trim(l_planaccion) & "'," & cambiafecha(l_planfecharev,"YMD",false) & ","&l_evldrnro &")"
else
	if l_tipo = "M" then
		l_sql = "UPDATE evaplan SET "
		l_sql = l_sql & " aspectomejorar = '" & trim(l_aspectomejorar) & "',"
		l_sql = l_sql & " planaccion     = '" & trim(l_planaccion) & "',"
		l_sql = l_sql & " planfecharev   = " & cambiafecha(l_planfecharev,"YMD",false) & ""
		l_sql = l_sql & " WHERE evaplan.evaplnro = "  & l_evaplnro
	else
		l_sql = "DELETE from evaplan where evaplan.evaplnro = "  & l_evaplnro
	end if	
end if	
set l_cm = Server.CreateObject("ADODB.Command")  
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.close
set cn=nothing

response.write "<script> parent.location.reload(); </script>"
%>
