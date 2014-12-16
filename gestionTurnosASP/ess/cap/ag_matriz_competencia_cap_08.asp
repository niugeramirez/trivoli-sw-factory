<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--
Archivo: ag_matriz_competencia_cap_08.asp
Descripción: Eliminar.
Autor : Lisandro Moro
Fecha: 29/03/2004
Modificado: 
-->
<% 
Dim l_cm
Dim l_rs
Dim l_sql

Dim l_evafacnro
Dim l_origen1
Dim l_origen2

l_evafacnro = request.querystring("cabnro")
l_origen1   = request.querystring("origen1")
l_origen2   = request.querystring("origen2")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT entnro, origen1, origen2"
l_sql = l_sql & " FROM cap_capacita"
l_sql = l_sql & " WHERE origen1 =" & l_origen1 & " AND entnro=" & l_evafacnro
l_sql = l_sql & " AND origen2 =" & l_origen2

rsOpen l_rs, cn, l_sql, 0

if not l_rs.eof then
   if l_origen1 <> "5" and l_origen2 <> "3" then
      Response.write "<script>alert('No se puede Modificar la Competencia ya que no fue cargada de forma Manual.');window.close();</script>"

      else	l_rs.close

		    set l_cm = Server.CreateObject("ADODB.Command")

		    l_sql = "DELETE FROM cap_capacita "
			l_sql = l_sql & " WHERE origen1 =" & l_origen1 & " AND entnro=" & l_evafacnro
            l_sql = l_sql & " AND origen2 =" & l_origen2

		    l_cm.activeconnection = Cn
		    l_cm.CommandText = l_sql
	 	    cmExecute l_cm, l_sql, 0
	end if	
end if	
	
cn.Close
Set cn = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location.reload();window.close();</script>"

%>
