<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--#include virtual="/turnos/shared/inc/sqls.inc"-->
<% 
'Archivo: listadeprecios_con_03.asp
'Descripción: ABM de Lista de Precios
'Autor : Raul Chinestra
'Fecha: 01/07/2015

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

Dim l_id
Dim l_titulo
Dim l_fecha
Dim l_flag_activo
Dim l_idobrasocial
Dim l_lpcab
Dim l_nueva_lpcab

l_tipo 		  = request.Form("tipo")
l_id 	      = request.Form("id")
l_titulo	  = request.Form("titulo")
l_fecha       = request.Form("fecha")
l_flag_activo = request.Form("activo")
l_idobrasocial = request.Form("idobrasocial")
l_lpcab		   = request.Form("lpcab")

' ------------------------------------------------------------------------------------------------------------------
' codigogenerado() :
' ------------------------------------------------------------------------------------------------------------------
function codigogenerado(tabla)
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("next_id",tabla)
	rsOpen l_rs, cn, l_sql, 0
	codigogenerado=l_rs("next_id")
	l_rs.Close
	Set l_rs = Nothing
end function 'codigogenerado()

Set l_rs = Server.CreateObject("ADODB.RecordSet")

if len(l_fecha) = 0 then
	l_fecha = "null"
else 
	l_fecha = cambiafecha(l_fecha,"YMD",true)	
end if 

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "INSERT INTO listaprecioscabecera "
	l_sql = l_sql & " (titulo, fecha, idobrasocial, flag_activo, idpreciocabeceraorigen ,empnro,created_by,creation_date,last_updated_by,last_update_date)"
	l_sql = l_sql & " VALUES ('" & l_titulo & "'," & l_fecha & "," & l_idobrasocial & "," & l_flag_activo &"," & l_lpcab & "," & session("empnro") &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
	
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	
	if l_lpcab <> "" then
		'Copio la lista de precios recibida
		l_nueva_lpcab = codigogenerado("listaprecioscabecera")	
		
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM listapreciosdetalle "
		l_sql = l_sql & " WHERE idlistaprecioscabecera= " & l_lpcab

		rsOpen l_rs, cn, l_sql, 0
		do while not l_rs.eof 	
		
			l_sql = "INSERT INTO listapreciosdetalle "
			l_sql = l_sql & " (idpractica, precio, idlistaprecioscabecera ,empnro,created_by,creation_date,last_updated_by,last_update_date)"
			l_sql = l_sql & " VALUES (" & l_rs("idpractica") & "," &Replace(l_rs("precio"), ",", ".")   & "," & l_nueva_lpcab &"," & session("empnro") &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"	
		
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			
			l_rs.movenext	
		loop
		l_rs.close	
	end if 'copiar lista
	
else
	l_sql = "UPDATE listaprecioscabecera "
	l_sql = l_sql & " SET titulo = '" & l_titulo & "'"
	l_sql = l_sql & " , fecha = " & l_fecha 
	'l_sql = l_sql & " , idobrasocial = " & l_idobrasocial 'Eugenio 03/04/2015 esto para mi no va, se me abortaba al agregar los campos who, creo que nunca se testeo
	l_sql = l_sql & " , flag_activo = " & l_flag_activo
	l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
	l_sql = l_sql & "    ,last_update_date = GETDATE()" 	
	l_sql = l_sql & " WHERE id = " & l_id
	
	'response.write l_sql & "<br>"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	

	if l_lpcab <> "" then
		'Copio la lista de precios recibida
		l_nueva_lpcab = codigogenerado("listaprecioscabecera")	
		
		l_sql = "SELECT * "
		l_sql = l_sql & " FROM listapreciosdetalle "
		l_sql = l_sql & " WHERE idlistaprecioscabecera= " & l_lpcab

		rsOpen l_rs, cn, l_sql, 0
		do while not l_rs.eof 	
		
			l_sql = "INSERT INTO listapreciosdetalle "
			l_sql = l_sql & " (idpractica, precio, idlistaprecioscabecera ,empnro,created_by,creation_date,last_updated_by,last_update_date)"
			l_sql = l_sql & " VALUES (" & l_rs("idpractica") & "," & Replace(l_rs("precio"), ",", ".")  & "," & l_id &"," & session("empnro") &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"	
		
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			
			l_rs.movenext	
		loop
		l_rs.close	
	end if 'copiar lista
	
end if

Set l_cm = Nothing

Response.write "OK"
%>

