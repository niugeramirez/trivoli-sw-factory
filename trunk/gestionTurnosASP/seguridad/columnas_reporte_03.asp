<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: columnas_reporte_03.asp
Descripcion: Modulo que se encarga de las altas/modificaciones
             de las columnas del confrep.
Modificacion:
    29/07/2003 - Scarpa D. - Agregado de la columna confsuma   
    13/08/2003 - Scarpa D. - Copiar la etiqueta de una columna en todas
	                         las que tienen el mismo numero.   	
    25/08/2003 - Scarpa D. - Agregado de la columna confval2								 
-----------------------------------------------------------------------------
-->

<% 
' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql
dim l_rs

Dim l_repnro
Dim l_confnrocol
Dim l_confetiq
Dim l_conftipo
Dim	l_confval
Dim	l_confval2
Dim	l_confaccion

Dim l_confnrocolant
Dim l_confetiqant
Dim l_conftipoant
Dim	l_confvalant
Dim	l_confval2ant
Dim	l_confaccionant

' traer valores del form de alta/modificacion -------------------------------------

l_tipo = request("tipo")

l_repnro	  = Request.form("repnro")

l_confnrocol  = Request.form("confnrocol")
l_confetiq	  = Request.form("confetiq")
l_conftipo    = Request.form("conftipo")
l_confval     = Request.form("confval")
l_confval2    = Request.form("confval2")
l_confaccion  = Request.form("confaccion")

l_confnrocolant  = Request.form("confnrocolant")
l_confetiqant    = Request.form("confetiqant")
l_conftipoant    = Request.form("conftipoant")
l_confvalant     = Request.form("confvalant")
l_confval2ant    = Request.form("confval2ant")
l_confaccionant  = Request.form("confaccionant")

set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 

		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT * FROM  confrep"
		l_sql = l_sql & " WHERE confrep.repnro		= "   & l_repnro
		l_sql = l_sql & " AND   confrep.confnrocol	= "   & l_confnrocol
		l_sql = l_sql & " AND   confrep.conftipo	= '"  & l_conftipo & "'"
		l_sql = l_sql & " AND   confrep.confetiq	= '"  & l_confetiq & "'"
		l_sql = l_sql & " AND   confrep.confval		= "   & l_confval
		l_sql = l_sql & " AND   confrep.confval2	= '"  & l_confval2	& "'"
		l_sql = l_sql & " AND   confrep.confaccion		= '"  & l_confaccion & "'"  		
		
		l_rs.MaxRecords = 1
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			Response.write "<script>alert('Ya existe otra columna con estos valores.');window.close();</script>"		
		else
			l_sql = "INSERT INTO confrep "
			l_sql = l_sql & "(conftipo , confetiq, confval, confval2, confnrocol, repnro, empnro, confaccion) "
			l_sql = l_sql & " values ('" 
			l_sql = l_sql & l_conftipo   & "','" 
			l_sql = l_sql & l_confetiq   & "'," 
			l_sql = l_sql & l_confval  & ",'" 
			l_sql = l_sql & l_confval2 & "'," 			
			l_sql = l_sql & l_confnrocol & "," 
			l_sql = l_sql & l_repnro     & ",1,'"
			l_sql = l_sql & l_confaccion & "')"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0	
		end if	
else
		if (l_confnrocolant <> l_confnrocol) or _  
		   (l_conftipoant   <> l_conftipo  ) or _
		   (l_confetiqant   <> l_confetiq  ) or _
		   (l_confvalant    <> l_confval   ) or _
		   (l_confvalant2   <> l_confval2  ) or _		   
		   (l_confaccionant <> l_confaccion)   then
		   ' algo modifico asi que me fijo si existe otrto reg. con lo nuevo
			Set l_rs = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT * FROM  confrep"
			l_sql = l_sql & " WHERE confrep.repnro		= "   & l_repnro
			l_sql = l_sql & " AND   confrep.confnrocol	= "   & l_confnrocol
			l_sql = l_sql & " AND   confrep.conftipo	= '"  & l_conftipo & "'"
			l_sql = l_sql & " AND   confrep.confetiq	= '"  & l_confetiq & "'"
			l_sql = l_sql & " AND   confrep.confval		= "   & l_confval
			l_sql = l_sql & " AND   confrep.confval2	= '"  & l_confval2 & "'"			
			l_sql = l_sql & " AND   confrep.confaccion	= '"   & l_confaccion & "'"			
			l_rs.MaxRecords = 1
			'response.write l_sql			
			rsOpen l_rs, cn, l_sql, 0 

			if not l_rs.eof then
				Response.write "<script>alert('Ya existe otra columna con estos valores.');window.close();</script>"		
			else
				l_sql = "UPDATE confrep SET "
				l_sql = l_sql & " conftipo		= '" & l_conftipo   & "', " 
				l_sql = l_sql & " confetiq		= '" & l_confetiq   & "', " 
				l_sql = l_sql & " confval		= "  & l_confval    & " , " 
				l_sql = l_sql & " confval2   	= '" & l_confval2   & "', " 
				l_sql = l_sql & " confnrocol	= "  & l_confnrocol & " , " 
				l_sql = l_sql & " confaccion	= '"  & l_confaccion & "'"		
				l_sql = l_sql & " WHERE repnro  = " & l_repnro
				l_sql = l_sql & " AND   confnrocol  = "   & l_confnrocolant
				l_sql = l_sql & " AND   conftipo	= '"  & l_conftipoant & "'"
				l_sql = l_sql & " AND   confetiq	= '"  & l_confetiqant & "'"
				l_sql = l_sql & " AND   confval		= "   & l_confvalant
				l_sql = l_sql & " AND   confval2	= '"  & l_confval2ant & "'"	
				l_sql = l_sql & " AND   confaccion	= '"   & l_confaccionant & "'"
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				'response.write l_sql
				cmExecute l_cm, l_sql, 0	
				
                'Copio la etiqueta a todas las columnas que tiene el mismo numero
				l_sql = "UPDATE confrep SET "
				l_sql = l_sql & " confetiq		= '" & l_confetiq   & "' " 
				l_sql = l_sql & " WHERE repnro  = " & l_repnro
				l_sql = l_sql & " AND   confnrocol  = "   & l_confnrocol
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				'response.write l_sql
				cmExecute l_cm, l_sql, 0	
				
			end if
		end if ' modifico algo	
end if

Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'columnas_reporte_01.asp?repnro=" & l_repnro & "';window.close();</script>"
%>
