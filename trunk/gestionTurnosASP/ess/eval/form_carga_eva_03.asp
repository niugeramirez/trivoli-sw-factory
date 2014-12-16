<% Option Explicit %>

<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'=======================================================================================
'Archivo	: form_carga_eva_03.asp
'Autor		: ?
'Fecha		: ?
'Modificacion: 03-06-2004 - CCRossi - adaptarlo para los nuevos campos de evacab y evadetevldor
'Modificacion: 01-12-2004-CCRossi -controlar que existe la asp de bussqueda del ternro del Rol
'Modificacion: 07-12-2004-CCRossi -cambiar tipo de cursor para recorrer recorset.
'----------------------------------------------------------------------------------
Dim objOpenFile, objFSO, strPath

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs
Dim l_rs2
Dim l_rs3
Dim l_rs_oblig
Dim l_rs_secc

Dim l_evacabnro
Dim l_evatevnro
Dim l_evldrnro
Dim l_evaluador
Dim l_evaseccnro
Dim l_evaetanro
Dim l_evarolaspdet 
Dim l_habilitado   
Dim l_evaseccmail

Dim l_fechahab
Dim l_horahab 
Dim l_hora 
Dim l_arrhr
		
'parametros
Dim l_empleado
Dim l_revisor
Dim l_evaevenro

l_empleado  = request.querystring("ternro")
l_evaevenro = request.querystring("evaevenro")

function strto2(cad)
	if trim(cad) <>"" then
		if len(cad)<2 then
			strto2= "0" & cad
		else
			strto2= cad
		end if 
	else
		strto2= "00"
	end if	
end function

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaforeta.evaetanro "
l_sql = l_sql & " FROM evaforeta "
l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evatipnro= evaforeta.evatipnro"
l_sql = l_sql & " WHERE evaforeta.evadef = -1"
l_sql = l_sql & " AND   evaevento.evaevenro = " & l_evaevenro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then
	l_evaetanro = l_rs("evaetanro")
else
	l_evaetanro = "null"
end if		
l_rs.Close
set l_rs=nothing
	
cn.BeginTrans

set l_cm = Server.CreateObject("ADODB.Command")
l_sql = "INSERT INTO evacab "
l_sql = l_sql & " (evaevenro , empleado, cabevaluada, cabaprobada "
if not isNull(l_evaetanro) then
l_sql = l_sql & ",evaetanro"
end if
l_sql = l_sql & ") "
l_sql = l_sql & " VALUES (" & l_evaevenro & ", " & l_empleado & ", 0, 0" 
if not isNull(l_evaetanro) then
l_sql = l_sql & "," & l_evaetanro 
end if
l_sql = l_sql & ")"
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
		

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = fsql_seqvalue("id","evacab")
rsOpen l_rs, cn, l_sql, 0
l_evacabnro=l_rs("id")
l_rs.Close
set l_rs=nothing

'Crear los EvaDet ...........................................................
Set l_rs_secc = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT evaseccnro "
l_sql = l_sql & " FROM evasecc "
l_sql = l_sql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro= evasecc.evatipnro"
l_sql = l_sql & " INNER JOIN evaevento  ON evaevento.evatipnro = evatipoeva.evatipnro "
l_sql = l_sql & " WHERE evaevenro = " & l_evaevenro
'rsOpen l_rs_secc, cn, l_sql, 0
rsOpenCursor l_rs_secc, cn, l_sql, 0, 3
do until l_rs_secc.eof
	l_evaseccnro = l_rs_secc("evaseccnro")

	l_sql = "INSERT INTO evadet "
	l_sql = l_sql & " (evacabnro , evaseccnro, detcargada) "
	l_sql = l_sql & " VALUES (" & l_evacabnro & ", " & l_evaseccnro & ", 0)"
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
			
	' Crear los evadetevldor ...............................................
	Set l_rs_oblig = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evaoblieva.evatevnro, evatevobli, evarolaspdet, afteranterior, evaobliorden, evaseccmail "
	l_sql = l_sql & " FROM evaoblieva "
	l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evaoblieva.evatevnro "
	l_sql = l_sql & " INNER JOIN evasecc      ON evasecc.evaseccnro= evaoblieva.evaseccnro "
	l_sql = l_sql & " LEFT  JOIN evarolasp ON evarolasp.evarolnro = evatipevalua.evarolnro "
	l_sql = l_sql & " WHERE evaoblieva.evaseccnro = " & l_evaseccnro
	l_sql = l_sql & " ORDER BY evaobliorden " 
	'rsOpen l_rs_oblig, cn, l_sql, 0
	rsOpenCursor l_rs_oblig, cn, l_sql, 0, 3
	do until l_rs_oblig.eof

		l_evatevnro	   = l_rs_oblig("evatevnro")
		l_evarolaspdet = l_rs_oblig("evarolaspdet") 'ASP que busca el ternro del evaluador
		l_habilitado   = not l_rs_oblig("afteranterior")
		l_evaseccmail  = l_rs_oblig("evaseccmail")
		
		l_evaluador= "null"
		l_fechahab="null"
		l_horahab =""
		if l_habilitado=-1 then
			l_hora = mid(time,1,8)
			l_arrhr= Split(l_hora,":")
			l_hora = strto2(l_arrhr(0)) & l_arrhr(1)	
			l_fechahab = cambiafecha(Date(),"","") 
			l_horahab  =  l_hora
		end if
			
		l_sql = "INSERT INTO evadetevldor "
		l_sql = l_sql & "(evacabnro , evaseccnro, evatevnro, evaluador, evldorcargada,habilitado,fechahab,horahab) "
		l_sql = l_sql & " VALUES (" & l_evacabnro & ", " & l_evaseccnro & ",  " & l_evatevnro & ",  " & l_evaluador & ", 0,"&l_habilitado&","&l_fechahab &",'"&l_horahab&"')"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = fsql_seqvalue("evldrnro","evadetevldor")
		rsOpen l_rs, cn, l_sql, 0
		l_evldrnro=l_rs("evldrnro")
		l_rs.Close
		Set l_rs = Nothing
		
		
		' se pierde el evldrnro....
		Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evldrnro, habilitado "
		l_sql = l_sql & " FROM evadetevldor "
		l_sql = l_sql & " WHERE evadetevldor.evacabnro = " & l_evacabnro
		l_sql = l_sql & " AND   evadetevldor.evatevnro = " & l_evatevnro
		l_sql = l_sql & " AND   evadetevldor.evaseccnro = " & l_evaseccnro
		rsOpen l_rs3, cn, l_sql, 0
		if not l_rs3.eof then
			l_evldrnro=l_rs3("evldrnro")
			l_habilitado = l_rs3("habilitado")
		end if	
		l_rs3.Close
		set l_rs3=nothing
			
		' Buscar el ternro para el evaluador.................................
		if trim(l_evarolaspdet) <>"" and trim(l_evldrnro)<>"" and (trim(l_evaluador)="null" or trim(l_evaluador)="" or isnull(l_evaluador)) then
			strPath = Server.MapPath(l_evarolaspdet)
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			If not objFSO.FileExists(strPath) Then%>
			<script>
			alert('El programa de Búsqueda de Evaluador no existe.\nVerifique la existencia de <%=l_evarolaspdet%>');
			</script>
			<%else%>
			<script>
			var r = showModalDialog('<%=l_evarolaspdet%>?ternro=<%=l_empleado%>&evldrnro=<%=l_evldrnro%>', '','dialogWidth:40;dialogHeight:22');
			</script>
			<%end if
		end if	
		
		if l_habilitado=-1 and cUsaMail= -1 and l_evaseccmail=-1 then%>
			<script>
			var r = showModalDialog('enviomail_eva_00.asp?evldrnro=<%=l_evldrnro%>', '','dialogWidth:40;dialogHeight:22');
			</script>
		<%end if
		
		l_rs_oblig.MoveNext
	loop
	l_rs_oblig.Close
	set l_rs_oblig=nothing
				
	l_rs_secc.MoveNext

loop
l_rs_secc.Close
set l_rs_secc=nothing

cn.CommitTrans

Set l_cm = Nothing	
%>
<script>
var r = showModalDialog('tomarrevisor_eva_00.asp?evacabnro=<%=l_evacabnro%>', '','dialogWidth:40;dialogHeight:22');
if (r!==0)				
{
	parent.document.datos.rempleg.value=r;
	parent.buscarrevisor();
}	
parent.cargardatos();
</script>

