<% Option Explicit %>

<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<% 
on error goto 0
'Archivo: asignar_cascara_con_07.asp
'Descripción: Asignar la patente y el chasis de acuerdo al transportista seleccionado
'Autor : Raul CHinestra
'Fecha: 11/05/2005

Dim l_camnro
Dim l_camcha
Dim l_camaco
Dim l_tranro

Dim l_cant
'Dim l_celdenro

Dim l_stkactor
Dim l_stkactde

Dim l_rs
Dim l_sql

l_camnro    = request.QueryString("camnro")


'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

	l_sql = "SELECT * "
	l_sql = l_sql & "FROM tkt_camionero "
	l_sql = l_sql & "WHERE camnro =" & l_camnro 

	rsOpen l_rs, cn, l_sql, 0

	if  (l_rs.eof) then	
		l_camcha = ""
		l_camaco = ""
		l_tranro = 0

		l_rs.close	
	else
		l_camcha = l_rs("camcha")
		l_camaco = l_rs("camaco")
		
		l_rs.close	
	
		l_sql = "SELECT tranro "
		l_sql = l_sql & "FROM tkt_cam_tra "
		l_sql = l_sql & " WHERE camnro =" & l_camnro
		rsOpen l_rs, cn, l_sql, 0
		
		l_cant = 0
		l_tranro = 0
		Do while not l_rs.eof
			l_tranro = l_Rs("tranro")
			l_cant = l_cant + 1
			l_rs.movenext
		loop
		if l_cant > 1 then 
			l_tranro = 0   ' Si tiene mas que una empresa transportista asociada que no seleccione ninguna de ellas
		end if	
		l_rs.close
	end if


Set l_rs = nothing

%>
<script>
	parent.actualizar_datos('<% =l_camcha %>','<% =l_camaco %>','<% =l_tranro %>')
</script>
