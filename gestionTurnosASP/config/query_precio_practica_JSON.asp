<% Option Explicit %>

<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<% 

Dim l_practicaid


Dim l_precio

Dim l_idos


Dim l_rs
Dim l_sql

l_practicaid = request("practicaid")

l_idos = request("idos")

'Response.write "l_idos "&l_idos

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

	if l_idos = "" then
		l_precio = 0
	else
		
		l_sql = "SELECT precio "
		l_sql = l_sql & " FROM listaprecioscabecera "
		l_sql = l_sql & " INNER JOIN listapreciosdetalle ON listapreciosdetalle.idlistaprecioscabecera = listaprecioscabecera.id "
		l_sql = l_sql & " WHERE flag_activo = -1 " 
		l_sql = l_sql & " AND idobrasocial = " & l_idos
		l_sql = l_sql & " AND idpractica = " & l_practicaid
		l_sql = l_sql & " and listaprecioscabecera.empnro = " & Session("empnro")
		'response.write l_sql
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_precio = l_rs("precio")
		else
			l_precio = 0
		end if
		l_rs.Close
		
	end if
	

Response.write "[{""resultado"":""OK"",""precio"":""" & l_precio & """}]"

%>

