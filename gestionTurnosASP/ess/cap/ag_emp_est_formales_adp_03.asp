<%option explicit%>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
'Archivo: emp_est_formales_adp_03.asp
'Descripción: abm de estudios formales - Grabar informacion
'Autor : Lisandro Moro
'Fecha: 11/09/2003
'Modificado: 03-10-2003 - CCRossi - Modificar chequeo de si exise el 
'									registro en la modificacion
'Modificado: 06-10-2003 - CCRossi - Sacar carrera si es completo
'Modificado: 07-10-2003 - CCRossi - Poner valoores que vienen vacios o 0 en NULL
'					para los campos que son parte de la clave de acceso
'Modificado 25-02-2004 - Scarpa D. - Se agrego la opcion de estudio actual
'Modificado 25-02-2004 - Scarpa D. - Se agrego el campo descripcion
'			17-10-2005 - Leticia A. - Adaptacion a Autogestion - arreglo el Alta de Est. F
'-------------------------------------------------------------------------------
%>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<%
on error goto 0

' de base de datos
dim l_rs
dim l_sql
Dim l_cm

Dim l_tipo
Dim l_ternro
Dim l_titnro
Dim l_instnro
Dim l_carredunro

Dim l_nivnroant
Dim l_titnroant
Dim l_instnroant
Dim l_carredunroant

Dim l_nivnro
Dim l_capcomp
Dim l_capcanmat
Dim l_capestact
Dim l_capanocur
Dim l_capfecdes
Dim l_capfechas
Dim l_caprango
Dim l_capprom
Dim l_capactual
Dim l_futdesc

'ternro, nivnro, titnro, instnro, carredunro, capcomp, capcanmat, capestact, capanocur, capfecdes, capfechas, caprango, capprom
'l_ternro  & ", " & l_nivest & ", " & l_titulo & ", " & l_institucion & ", " & l_carrera & ", "  & l_capcomp & ", " & l_capcantmat & ", " & l_capestact & ", " & l_capanocur & ", " & l_capfecdes & ", " & l_capfechas & ", " & l_caprango & ", " & l_capprom

l_tipo			= Request.Form("tipo")
l_ternro 		= l_ess_ternro
l_titnro 		= Request.Form("titnro")
l_instnro	 	= Request.Form("instnroaux")
l_carredunro	= Request.Form("carredunro")
l_nivnro 		= Request.Form("nivnro")

'response.write(l_instnro) & "<br>"

l_titnroant		= Request.Form("titnroant")
l_instnroant 	= Request.Form("instnroant")
l_carredunroant	= Request.Form("carredunroant")
l_nivnroant		= Request.Form("nivnroant")

l_capcomp 		= Request.Form("capcomp")
l_capcanmat 	= Request.Form("capcanmat")
l_capestact 	= Request.Form("capestact")
l_capanocur 	= Request.Form("capanocur")
l_capfecdes 	= Request.Form("capfecdes")
l_capfechas 	= Request.Form("capfechas")
l_caprango 		= Request.Form("caprango")
l_capprom 		= Request.Form("capprom")

if l_titnro = "" or l_titnro = "0" then
	l_titnro = "null"
end if
if l_titnroant = "" or l_titnroant = "0" then
	l_titnroant = "null"
end if

if l_instnro = "" or l_instnro = "0" then
	l_instnro = "null"
end if
if l_instnroant = "" or l_instnroant = "0" then
	l_instnroant = "null"
end if

if l_carredunro = "" or l_carredunro = "0" then
	l_carredunro = "null"
end if
if l_carredunroant = "" or  l_carredunroant = "0" then
	l_carredunroant = "null"
end if


if l_capcomp = "" then
	l_capcomp = "0"
end if
if l_capestact = "" then
	l_capestact = "0"
end if

'if l_capcomp = "-1" then
	'l_capcomp = 1
'end if
'if l_capestact = "-1" then
	'l_capestact = 1
'end if

if l_capcanmat = "" then
	l_capcanmat = "null"
end if
if l_capanocur = "" then
	l_capanocur = "null"
end if
if l_caprango = "" then
	l_caprango = null
end if
if l_capprom = "" then
	l_capprom = null
end if
if not IsDate(l_capfecdes) then
	l_capfecdes = "null"
else
	l_capfecdes = cambiafecha(l_capfecdes,"YMD",true)
end if

if not IsDate(l_capfechas) then
	l_capfechas = "null"
else
	l_capfechas = cambiafecha(l_capfechas,"YMD",true)
end if


'response.write(l_carredunro) & "<br>"


set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	'Valido que no hayan registro duplicados por key
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT   ternro, nivnro, titnro, instnro, carredunro"
	l_sql = l_sql & " FROM  cap_estformal "
	l_sql = l_sql & " WHERE ternro = " & l_ternro 
	if l_titnro ="null" then
		l_sql = l_sql & " AND titnro IS NULL "
	else
		l_sql = l_sql & " AND titnro = " & l_titnro 
	end if
	if l_instnro ="null" then
		l_sql = l_sql & " AND instnro IS NULL " 
	else
		l_sql = l_sql & " AND instnro = " & l_instnro 
	end if
	if l_carredunro ="null" then
		l_sql = l_sql & " AND carredunro IS NULL "
	else
		l_sql = l_sql & " AND carredunro = " & l_carredunro
	end if
	l_sql = l_sql & " AND nivnro = " & l_nivnro 
	
	rsOpen l_rs, cn, l_sql, 0 
	
'	response.write l_sql & "<br>"
	
	if not l_rs.eof then
		Response.write "<script>alert('Ya existe un Estudio Formal con las mismas características, por favor modifique los datos.');window.close();</script>"
	else
		l_sql = "INSERT INTO cap_estformal "
		l_sql = l_sql & "(ternro, nivnro, titnro, instnro, carredunro, capcomp, capcantmat, capestact, capanocur, capfecdes, capfechas, caprango, capprom)"
		l_sql = l_sql & " VALUES ( "
		l_sql = l_sql & " "& l_ternro  & ", " & l_nivnro & ", " & l_titnro & ", " & l_instnro & ", " & l_carredunro & ", "  & l_capcomp & ", " & l_capcanmat & ", " & l_capestact & ", " & l_capanocur & ", " & l_capfecdes & ", " & l_capfechas & ", '" & l_caprango & "', '" & l_capprom & "'"
		l_sql = l_sql & " ) "
	end if
	l_rs.close
	set l_rs=nothing
	
else
	'Valido que no hayan registro duplicados por key
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT   ternro, nivnro, titnro, instnro, carredunro"
	l_sql = l_sql & " FROM  cap_estformal "
	l_sql = l_sql & " WHERE ternro = " & l_ternro 
	l_sql = l_sql & " AND (nivnro = " & l_nivnro
	if l_instnro ="null" then
		l_sql = l_sql & " AND instnro IS NULL " 
	else
		l_sql = l_sql & " AND instnro = " & l_instnro 
	end if
	if l_titnro ="null" then
		l_sql = l_sql & " AND titnro IS NULL "
	else
		l_sql = l_sql & " AND titnro = " & l_titnro 
	end if
	if l_carredunro ="null" then
		l_sql = l_sql & " AND carredunro IS NULL "
	else
		l_sql = l_sql & " AND carredunro = " & l_carredunro
	end if
	l_sql = l_sql & " ) AND NOT ( nivnro = " & l_nivnroant
	if l_instnroant ="null" then
		l_sql = l_sql & " AND instnro IS NULL" 
	else
		l_sql = l_sql & " AND instnro = " & l_instnroant
	end if
	if l_titnroant ="null" then
		l_sql = l_sql & " AND titnro IS NULL" 
	else
		l_sql = l_sql & " AND titnro = " & l_titnroant 
	end if
	if l_carredunroant="null" then
		l_sql = l_sql & " AND carredunro IS NULL "
	else
		l_sql = l_sql & " AND carredunro = " & l_carredunroant
	end if
	l_sql = l_sql & " ) " 
	rsOpen l_rs, cn, l_sql, 0 

'	response.write l_sql & "<br>"

	if not l_rs.eof then
		Response.write "<script>alert('Ya existe un Estudio Formal con las mismas características, por favor modifique los datos.');window.close();</script>"
	else
		l_sql = "UPDATE cap_estformal "
		l_sql = l_sql & " SET " 
		l_sql = l_sql & "   ternro = "	& l_ternro
		l_sql = l_sql & " , nivnro  = "	& l_nivnro
		l_sql = l_sql & " , titnro = "	& l_titnro
		l_sql = l_sql & " , instnro = "	& l_instnro
		l_sql = l_sql & " , carredunro = " & l_carredunro
		l_sql = l_sql & " , capcomp = " & l_capcomp
		l_sql = l_sql & " , capcantmat = " & l_capcanmat
		l_sql = l_sql & " , capestact = " & l_capestact
		l_sql = l_sql & " , capanocur = " & l_capanocur
		l_sql = l_sql & " , capfecdes = " & l_capfecdes
		l_sql = l_sql & " , capfechas = " & l_capfechas
		l_sql = l_sql & " , caprango = '" & l_caprango & "' "
		l_sql = l_sql & " , capprom = '" & l_capprom & "' "
		l_sql = l_sql & " WHERE  ternro  = " & l_ternro
		l_sql = l_sql & " AND    nivnro  = " & l_nivnroant
		if l_instnroant ="null" then
			l_sql = l_sql & " AND    instnro IS NULL " 
		else
			l_sql = l_sql & " AND    instnro = " & l_instnroant
		end if
		if l_titnroant ="null" then
			l_sql = l_sql & " AND    titnro  IS NULL " 
		else
			l_sql = l_sql & " AND    titnro  = " & l_titnroant
		end if
		if l_carredunroant="null" then
			l_sql = l_sql & " AND carredunro  IS NULL "
		else
			l_sql = l_sql & " AND carredunro  = " & l_carredunroant
		end if	
	end if
	l_rs.close
	set l_rs=nothing
end if

'response.write l_sql & "<br>"

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

Set cn = Nothing
Set l_cm = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.location.reload();window.close();</script>"

%>
