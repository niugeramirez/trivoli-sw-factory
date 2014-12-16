<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<% 
'=====================================================================================
'Archivo  : grabar_estructuras_eva_00.asp
'Objetivo : grabar de estructuras (evaluaestr y evaestructuras)
'Fecha	  : 17-09-2004
'Autor	  : CCRossi
'Modificacion : 
'=====================================================================================

' variables
' parametros de entrada ----------------------------------------
  Dim l_evldrnro
  Dim l_tenro
  Dim l_estrnro
  Dim l_fdesde
  Dim l_fhasta
  Dim l_evaestrdext
  Dim l_tipo
    
' variables de base de datos ------------------------------------
  Dim l_cm
  Dim l_sql
  Dim l_rs
  Dim l_rs1
    
' locales
  Dim l_evacabnro 
  
' parametros de entrada
  l_evaestrdext  = left(trim(request.querystring("evaestrdext")),255)
  l_estrnro  = request.querystring("estrnro")
  l_evldrnro = request.querystring("evldrnro")
  l_tenro	 = request.querystring("tenro")
  l_fdesde	 = request.querystring("fdesde")
  l_fhasta	 = request.querystring("fhasta")
  l_tipo	 = request.querystring("tipo")

'----------------------------------------------------------
'       BODY 
'----------------------------------------------------------

'Response.Write("<script>alert('"&l_estrnro&"');</script>")

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evacabnro FROM evadetevldor "
l_sql = l_sql & " WHERE  evldrnro  = " & l_evldrnro
rsOpen l_rs1, cn, l_sql, 0
if not l_rs1.eof then
	l_evacabnro = l_rs1("evacabnro")
end if
l_rs1.Close
set l_rs1=nothing

select case l_tipo
case "M":
		l_sql = "UPDATE evaestructuras SET "
		l_sql = l_sql & " fdesde  = " & cambiafecha(l_fdesde,"","") & ","
		l_sql = l_sql & " fhasta  = " & cambiafecha(l_fhasta,"","") & ","
		l_sql = l_sql & " estrnro = " & l_estrnro 
		l_sql = l_sql & " WHERE tenro	  = "  & l_tenro
		l_sql = l_sql & " AND   evacabnro = "  & l_evacabnro
		set l_cm = Server.CreateObject("ADODB.Command") 
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0

		l_sql = "UPDATE evaluaestr SET "
		l_sql = l_sql & " evaestrdext = '" & trim(l_evaestrdext) & "',"
		l_sql = l_sql & " estrnro     = " & l_estrnro 
		l_sql = l_sql & " WHERE evldrnro = "  & l_evldrnro
		l_sql = l_sql & " AND   tenro    = "  & l_tenro
		set l_cm = Server.CreateObject("ADODB.Command") 
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0

case "B":
		l_sql = "DELETE FROM evaluaestr  "
		l_sql = l_sql & " WHERE evldrnro= " & l_evldrnro
		l_sql = l_sql & " AND   tenro   = " & l_tenro
		set l_cm = Server.CreateObject("ADODB.Command") 
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0

		l_sql = "DELETE FROM evaestructuras "
		l_sql = l_sql & " WHERE tenro     =" & l_tenro
		l_sql = l_sql & " AND   evacabnro =" & l_evacabnro
		set l_cm = Server.CreateObject("ADODB.Command") 
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
end select


response.write "<script> parent.location.reload(); </script>"
%>
