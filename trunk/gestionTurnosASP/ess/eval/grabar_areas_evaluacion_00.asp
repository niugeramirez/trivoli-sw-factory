<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
on error goto 0
' variables
' parametros de entrada ----------------------------------------
  Dim l_evatitnro
  Dim l_evldrnro
  Dim l_evaareadesc
  Dim l_evatrnro

  dim l_campo 
  
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
    
' parametros de entrada
  l_evatitnro	= Request.QueryString("evatitnro")
  l_evaareadesc = request.querystring("evaareadesc")

  l_evldrnro    = request.querystring("evldrnro")
  l_evatrnro    = request.querystring("evatrnro")

    l_campo = request.querystring("campo")
  
if len(trim(l_evaareadesc)) <> 0 then
   l_evaareadesc = left(trim(request.querystring("evaareadesc")),200)
end if 
if l_evatrnro="0" then
   l_evatrnro="null"
end if

'BODY ----------------------------------------------------------
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_sql = "UPDATE evaarea SET "
	l_sql = l_sql & " evaareadesc = '"   & l_evaareadesc & "',"
	l_sql = l_sql & " evatrnro    =  "		   & l_evatrnro 
	l_sql = l_sql & " WHERE evaarea.evatitnro = "  & l_evatitnro
	l_sql = l_sql & " AND   evaarea.evldrnro = "  & l_evldrnro
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	 ' response.write " <script> alert('"&l_campo&"')</script>"
	
	response.write "<script> parent.document.datos."&l_campo&".focus();</script>"
	Response.write " <script> window.close() </script>"
%>
