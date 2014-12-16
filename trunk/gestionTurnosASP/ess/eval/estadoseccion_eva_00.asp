<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'------------------------------------------------------------------------------------------------
'Archivo		: estadoseccion_eva_00.asp
'Autor			: CCRossi
'Fecha			: 12-11-2004
'Descripcion	: estado de la seccion para el evaluador
'idem estado_seccion_eva_AG pero desde el modulo con distinto look&feel
'-------------------------------------------------------------------------------------------------
on error goto 0
Dim l_rs
Dim l_rs1
Dim l_rs2
Dim l_sql

'locales
Dim l_evldorcargada 
Dim l_habilitado 
dim	l_evacabnro 
dim l_empleado 
dim l_cabaprobada
dim l_readonly
dim l_etaprogcarga
dim l_etaprogread 
dim l_depende ' guarda el nombre de la seccion de la cual depende esta y no esta TERMINADA
dim l_terminada ' habilita o no el tilde terminada

'parametros
dim l_evldrnro
dim l_evaseccnro
dim l_logeado ' 0: no; -1 si

l_evldrnro		= request("evldrnro")
l_evaseccnro	= request("evaseccnro")
l_cabaprobada	= request("cabaprobada")
l_logeado		= request("logeado")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evadetevldor.evacabnro, evadetevldor.evaseccnro, empleado, cabaprobada FROM evadetevldor INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro WHERE evldrnro=" & l_evldrnro
'l_sql = l_sql & " AND   evaseccnro=" & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0 
if not  l_rs.eof then
	l_evacabnro = l_rs("evacabnro")
	l_empleado = l_rs("empleado")
	l_evaseccnro= l_rs("evaseccnro")
	
	'if l_rs("cabaprobada")=-1 then
	'	l_cabaprobada = "checked disabled"
	'end if	
end if
l_rs.close	
set l_rs=nothing

'Chequear campo readonly (a ver si hablito Cab Evaluada )
'Para cada seccion obligatoria de la cabecera... 
' ver si hay al menos evldorcargada=0 no habilito Aprobada
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evadetevldor.evacabnro, empleado, cabaprobada, evaseceta.etaprogcarga, evaseceta.etaprogread FROM evadetevldor INNER JOIN evacab ON evacab.evacabnro= evadetevldor.evacabnro INNER JOIN evasecc ON evasecc.evaseccnro= evadetevldor.evaseccnro AND evasecc.evaoblig= -1"
l_sql = l_sql & " INNER JOIN evaoblieva ON evaoblieva.evaseccnro= evadetevldor.evaseccnro AND  evaoblieva.evatevnro= evadetevldor.evatevnro left  join evaseceta on evaseceta.evaseccnro= evadetevldor.evaseccnro AND  evaseceta.evatipnro= evasecc.evatipnro AND  evaseceta.evaetanro= evacab.evaetanro "
l_sql = l_sql & " WHERE evadetevldor.evacabnro = " & l_evacabnro & " AND evaoblieva.evatevobli = -1" & " AND evadetevldor.evldorcargada = 0" 
rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then
	l_readonly = ""
	l_cabaprobada =""
else	
	l_readonly = " readonly disabled"
	l_etaprogcarga = l_rs("etaprogcarga")
	l_etaprogread =  l_rs("etaprogread")
	if l_rs("cabaprobada")=-1 then
		l_cabaprobada = "checked disabled"
	end if	
end if
l_rs.close	
set l_rs=nothing

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>

<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function HabilitarProximo(){
	if (document.all.evldorcargada.checked)
	{
		document.all.evldorcargada.disabled=true;
		// actualizar edoldorcargada =-1 para el actual
		// Habilitar proximo evldrnro
		// Chequear Si estan todas las secciones obligatorias cargadas, habilitar campo aprobada
		var r = showModalDialog('habilitar_evaluador_eva_ag_00.asp?evldrnro=<%=l_evldrnro%>&evldorcargada=-1&evaseccnro=<%=l_evaseccnro%>', '','dialogWidth:20;dialogHeight:20'); 
		if (r==0){
			parent.actualizarcarga(<%=l_evldrnro%>,<%=l_evaseccnro%>,0,'<%=l_etaprogcarga%>','<%=l_etaprogread%>',0);	
			parent.actualizarevaluador(<%=l_evacabnro%>,<%=l_evaseccnro%>,<%=l_empleado%>);	
			}
		
	}	
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" width="96%">
<table id="tabla" width="95%">
<form name="datos" method="post">
<%
	if trim(l_evldrnro)="" then
	%>
    <tr>
        <td>Seleccione un Evaluador.</td>
    </tr>
<%
else
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  * FROM evadetevldor WHERE evldrnro=" & l_evldrnro & " AND   evaseccnro=" & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0 
do until l_rs.eof
	
	if l_rs("habilitado")=-1 then
		l_habilitado = "CHECKED"
	else	
		l_habilitado = ""
	end if	
	if l_rs("evldorcargada")=-1 then
		l_evldorcargada = "CHECKED"
	else	
		l_evldorcargada = ""
	end if	
	
	if l_rs("habilitado")=0 then ' si no está habilitado no permito tocar terminada
		l_terminada = "readonly disabled"
	else
		l_terminada=""
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT  depende.evaseccnro, depende.titulo dependenom FROM evasecc INNER JOIN evasecc depende ON depende.evaseccnro = evasecc.dependesecnro "
		l_sql = l_sql & " WHERE evasecc.evaseccnro = " & l_evaseccnro & " AND EXISTS ("
		l_sql = l_sql & " SELECT * FROM evacab INNER JOIN evadetevldor det ON det.evacabnro=evacab.evacabnro AND det.evldorcargada <> -1 AND det.evacabnro  = " & l_evacabnro 
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro AND evadetevldor.evaluador  = det.evaluador AND evadetevldor.evacabnro  = " & l_evacabnro & " AND evadetevldor.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " WHERE evacab.evacabnro = " & l_evacabnro  & " AND det.evaseccnro = depende.evaseccnro )"
		rsOpen l_rs1, cn, l_sql, 0
		'Response.Write l_sql
		if NOT l_rs1.eof then
			l_depende = l_rs1("dependenom")
			l_terminada = " readonly disabled"
			l_evldorcargada = "" ' le saco el tilde de terminada.. no debe estar terminada!
			
			l_rs1.Close
			set l_rs1 = nothing
		else
			l_rs1.Close
			set l_rs1 = nothing
			l_terminada=""
			l_depende  =""
			' verificar si la seccion es de objetivos y no se creo NINGUNO
			' no permitir terminarla
			Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT evaseccnro FROM evasecc INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro AND evatiposecc.tipsecobj =-1 WHERE evasecc.evaseccnro = " & l_evaseccnro & " AND EXISTS ("
			l_sql = l_sql & " SELECT * FROM evacab INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro AND evadetevldor.evacabnro  = " & l_evacabnro & " AND evadetevldor.evldrnro = " & l_evldrnro
			l_sql = l_sql & " INNER JOIN evaluaobj   ON evaluaobj.evldrnro = evadetevldor.evldrnro AND evaluaobj.evaborrador = 0 INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro=evaluaobj.evaobjnro WHERE evadetevldor.evaseccnro = evasecc.evaseccnro)"
			'Response.Write l_sql
			rsOpen l_rs1, cn, l_sql, 0
			if l_rs1.eof then
				Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
				l_sql = "SELECT evaseccnro FROM evasecc INNER JOIN evatiposecc ON evatiposecc.tipsecnro = evasecc.tipsecnro AND evatiposecc.tipsecobj =-1 WHERE evasecc.evaseccnro = " & l_evaseccnro
				'Response.Write l_sql
				rsOpen l_rs2, cn, l_sql, 0
				if l_rs2.EOF then
					l_terminada = ""
				else
					l_terminada = " readonly disabled"
				end if	
			else
				l_terminada = ""	
			end if
			l_rs1.Close
			set l_rs1 = nothing
		end if
		
	end if	

	%>
	<input type="Hidden" name="evldrnro" value ="<%=l_evldrnro%>">
    <tr>
		<td width="95%" align="center">
			<BR>
			<font size="1">H:
			<input type="Checkbox" <%= l_habilitado%>    style="width:20px;height:20px" disabled readonly name="habilitado"  > 
			<font size="1">T:
			<input readonly disabled type="Checkbox" <%= l_evldorcargada%> style="width:20px;height:20px" name="evldorcargada"> 
			<br>
        </td>
    </tr>
<%
l_rs.MoveNext
loop
l_rs.Close
set l_rs = Nothing
end if ' no habilitado
cn.Close
set cn = Nothing
%>
</table>

</form>
</body>
</html>

