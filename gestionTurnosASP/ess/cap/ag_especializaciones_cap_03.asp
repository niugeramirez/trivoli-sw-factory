<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_especializaciones_cap_03.asp
Descripcion: especializaciones
Autor: Lisandro Moro
Fecha: 29/03/2004
Modificado:
Martin Ferraro - 30/08/2007 - Correccion de tipos en ORACLE
-->
<% 
 on error goto 0
 
'ADO
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
' variables
 Dim l_grabar
 Dim l_ternro
 Dim l_eltananro
 Dim l_espnivnro
 Dim l_espestrrhh

 l_ternro   = l_ess_ternro
 l_grabar   = Request.QueryString("grabar")
 
 'l_lista= Split(l_grabar,";")		
'	 i = 1
'	 do while i <= UBound(l_lista)-1
'		response.write l_lista(i) & "<br>"
'		i = i + 1
'	 loop	
'	 response.end

'uso local
 dim i
 dim j
 dim l_lista
 dim l_datos
 
 'Al operar sobre varias tablas debo iniciar una transacción
 cn.BeginTrans
 
 'Creo los objetos ADO
 set l_rs = Server.CreateObject("ADODB.RecordSet")
 set l_cm = Server.CreateObject("ADODB.Command")
 
 'De la tabla relación busco los relacionados con la entidad y los borro
 l_sql = "SELECT eltananro, ternro, espnivnro, espmeses, espfecha "
 l_sql = l_sql & "  FROM especemp "
 l_sql = l_sql & " WHERE ternro = " & l_ternro
 rsOpenCursor l_rs, cn, l_sql,0,adOpenKeyset
 
 l_cm.activeconnection = Cn
 do until l_rs.eof
	l_sql = "DELETE FROM especemp "
	l_sql = l_sql & " WHERE ternro = " & l_ternro 
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
 	l_rs.MoveNext
 loop
 l_rs.close
 
 
	 l_lista= Split(l_grabar,";")		
	 i = 1
	 do while i <= UBound(l_lista)-1
	 	l_datos = Split(l_lista(i),",")
			j = 1
				do While j <= UBound(l_datos)-2
					l_eltananro	= l_datos(j)
					l_espnivnro	= l_datos(j+1)
					l_espestrrhh = l_datos(j+2)
					'response.write l_entnro & "-" & l_porcnro & "<BR>"
					l_sql = "INSERT INTO especemp "
					'l_sql = l_sql & "(eltananro, ternro, espnivnro, espmeses, espfecha) "
					'l_sql = l_sql & " VALUES (" & l_eltananro & ", " & l_ternro & " ," & l_espnivnro & ", NULL, " & cambiafecha(date(),"YMD",true) & " )"
					l_sql = l_sql & "(eltananro, ternro, espnivnro, espmeses, espfecha,espestrrhh) "
					l_sql = l_sql & " VALUES (" & l_eltananro & ", " & l_ternro & " ," & l_espnivnro & ", NULL, " & cambiafecha(date(),"YMD",true) & "," & l_espestrrhh & ")"
					l_cm.CommandText = l_sql
					cmExecute l_cm, l_sql, 0
					j = j + 3
				loop
			i = i + 1
	 loop	

 cn.CommitTrans
 
 Set cn = Nothing
 Set l_cm = Nothing

 response.write "<script>alert('Operación Realizada.');window.opener.opener.location.reload();window.opener.close();window.close();</script>"
%> 
