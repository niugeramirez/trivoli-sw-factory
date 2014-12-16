<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%Response.AddHeader "Content-Disposition", "attachment;filename=monitor_eventos.xls" %>
<%
'--------------------------------------------------------------------------
'Archivo	: monitor_evento_eva_excel.asp
'Descripción: salida excel del monitor de eventos
'Autor		: CCRossi
'Fecha		: 06-08-2004
'Modificado: 14-09-2004 - CCRossi agregar filtro por Detallado o Resumido
'Modificado: 15-09-2004 - CCRossi agregar filtro por Estructura
'--------------------------------------------------------------------------
'Variables base de datos
 Dim l_rs
 Dim l_sql

'Variables filtro y orden
 dim l_filtro
 dim l_filtro2
 dim l_orden
'parametros
 dim l_estado
 dim l_evaevenro
 dim l_obligatoria
 dim l_control    ' 0:No habilitados- 1:Habilitados y No Ingresaron-2:Ingresaron y No Terminaron-3:Todos
 dim l_mostrar    ' 0:Resumen - 1:Detallado
 dim l_listternro ' lista de ternro por estructura o empleado o evaluador.
   
'locales 
 dim l_empleg
 dim l_evatevnro
 dim l_evaseccnro
 dim l_evaoblig 
 dim l_cantidad
 dim l_evacabnro
 dim l_color
 dim l_letra
 dim l_cabaprobada    
 dim l_texto
 dim l_listamail
 dim l_terape	
 dim l_ternom	
 dim l_evaterape 
 dim l_evaternom 
 dim l_evatevdesabr	     

 dim l_fechaapro 
 dim l_ternro   
 
'suman eltotal de secciones de cada tipo, para el reporte resumen (total de no habilitados, etc)
 dim l_nohabilitado
 dim l_noingreso
 dim l_notermino

 				   
'Tomar parametros
 l_filtro = request("filtro")
 l_orden  = request("orden")
 l_filtro = request("filtro") 

 l_estado		= request("estado") 
 l_evaevenro	= request("evaevenro") 
 l_obligatoria	= request("evaoblig") 
 l_control		= request("control") 
 l_mostrar		= request("mostrar")  
 l_listternro	= request("listternro")  
 
 if trim(l_estado)="" then
	l_estado=-1
 end if
 if trim(l_obligatoria)=""  then
	l_obligatoria =2 ' ambas
 end if
 if trim(l_control)=""  then
	l_obligatoria =3 ' ambas
 end if
   
'Body 

 if l_orden = "" then
	l_orden = " empleado.empleg "
 end if
%>
<html>

<head>
<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title>Monitor - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <th nowrap><%if ccodelco=-1 then%>N&uacute;mero<%else%>Empleado<%end if%></th>
        <th nowrap>Apellido y Nombre</th>
        <th nowrap>Fecha Aprobaci&oacute;n</th>
        <th nowrap>Evaluador del Rol </th>
        <%if l_mostrar ="0" then ' DETALLADO%>	
        <th nowrap>Secci&oacute;n </th>
		<%if l_obligatoria = 2 or l_obligatoria = "" then%>
        <th nowrap>Secci&oacute;n Oblig.</th>
        <%end if%>
        <th nowrap>Fecha de Habilitaci&oacute;n </th>
        <th nowrap>Fecha de Ingreso </th>
        <th nowrap>Fecha de Finalizaci&oacute;n</th>
        <%else%>
        <th nowrap>No Habilitado</th>
        <th nowrap>No Ingresó</th>
        <th nowrap>No Terminó</th>
        <%end if%>
    </tr>

<%
dim eval(3)
dim auto(3)

For i = 0 To 2
	eval(i) = 0
Next
For i = 0 To 2
	auto(i) = 0
Next

if l_evaevenro <> "" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT DISTINCT evatipevalua.evatevnro,evatipevalua.evatevdesabr FROM empleado INNER JOIN evacab ON evacab.empleado = empleado.ternro AND evacab.evaevenro   ="  & l_evaevenro
	if cint(l_estado) <> 2 and l_estado <>"" then
	l_sql = l_sql & "		 AND evacab.cabaprobada ="  & l_estado
	end if
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro INNER JOIN evasecc	  ON evasecc.evaseccnro= evadetevldor.evaseccnro" 
	if cint(l_obligatoria) <>2 and l_obligatoria <>"" then
	l_sql = l_sql & " AND   evaoblig = " & l_obligatoria
	end if
	l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evadetevldor.evatevnro LEFT JOIN empleado evaluador ON evaluador.ternro= evadetevldor.evaluador WHERE evacab.evaevenro   = " & l_evaevenro
	if cint(l_estado) <> 2 and l_estado <>"" then
	l_sql = l_sql & " AND   evacab.cabaprobada = " & l_estado
	end if
	if trim(l_listternro)<>"" then
		l_sql = l_sql & " AND evacab.empleado IN (" & l_listternro & ")"
	end if	
	
	select case l_control
	case "0":
		l_sql = l_sql & " AND   (evadetevldor.habilitado=0 AND  evadetevldor.fechahab IS NULL AND evadetevldor.fechacar IS NULL)" 
		l_texto = " - No Habilitados "
	case "1":
		l_sql = l_sql & " AND   evadetevldor.habilitado=-1 AND evadetevldor.fechaing IS NULL AND evadetevldor.fechacar IS NULL" 
		l_texto = " - Habilitados y No Ingresaron "
	case "2":
		l_sql = l_sql & " AND   evadetevldor.habilitado=-1 AND evadetevldor.fechaing IS NOT NULL AND evadetevldor.fechacar IS NULL " 
		l_texto = " - Ingresaron y No Terminaron "
	case "3":
		l_sql = l_sql & " "
	end select	
	if trim(l_filtro)<>"" then
		l_sql = l_sql & "AND " & l_filtro
	end if	
	l_sql = l_sql & " ORDER BY evatipevalua.evatevnro,evatipevalua.evatevdesabr "
	
	dim leyendas(3)
	Dim i 
	i=0
	
	rsOpen l_rs, cn, l_sql, 0 
	DO WHILE not l_rs.eof
		leyendas(i) = l_rs("evatevdesabr")
		i=i+1
		l_rs.MoveNext
	loop
	l_rs.close
	set l_rs=nothing
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT evacab.evacabnro, empleado.ternro,empleado.empleg, empleado.terape, empleado.ternom, evacab.fechaapro, evacab.cabaprobada, evadetevldor.ingreso, evadetevldor.fechaing, evadetevldor.evldorcargada, evadetevldor.fechacar, evadetevldor.habilitado, evadetevldor.fechahab, "  
	l_sql = l_sql & " evasecc.evaseccnro,evasecc.evaoblig, evasecc.orden, evasecc.titulo, evatipevalua.evatevnro,evatipevalua.evatevdesabr, evaluador.empleg evaempleg, evaluador.ternro evaternro, evaluador.terape evaterape, evaluador.ternom evaternom  FROM  empleado " 
	l_sql = l_sql & " INNER JOIN evacab ON evacab.empleado = empleado.ternro AND evacab.evaevenro   ="  & l_evaevenro
	if cint(l_estado) <> 2 and l_estado <>"" then
	l_sql = l_sql & "		 AND evacab.cabaprobada ="  & l_estado
	end if
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro INNER JOIN evasecc	  ON evasecc.evaseccnro= evadetevldor.evaseccnro" 
	if cint(l_obligatoria) <>2 and l_obligatoria <>"" then
	l_sql = l_sql & " AND   evaoblig = " & l_obligatoria
	end if
	l_sql = l_sql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evadetevldor.evatevnro LEFT JOIN empleado evaluador ON evaluador.ternro= evadetevldor.evaluador WHERE evacab.evaevenro   = " & l_evaevenro
	if cint(l_estado) <> 2 and l_estado <>"" then
	l_sql = l_sql & " AND   evacab.cabaprobada = " & l_estado
	end if
	if trim(l_listternro)<>"" then
		l_sql = l_sql & " AND evacab.empleado IN (" & l_listternro & ")"
	end if	
	
	select case l_control
	case "0":
		l_sql = l_sql & " AND   (evadetevldor.habilitado=0 AND  evadetevldor.fechahab IS NULL AND evadetevldor.fechacar IS NULL)" 
		l_texto = " - No Habilitados "
	case "1":
		l_sql = l_sql & " AND   evadetevldor.habilitado=-1 AND evadetevldor.fechaing IS NULL AND evadetevldor.fechacar IS NULL" 
		l_texto = " - Habilitados y No Ingresaron "
	case "2":
		l_sql = l_sql & " AND   evadetevldor.habilitado=-1 AND evadetevldor.fechaing IS NOT NULL AND evadetevldor.fechacar IS NULL " 
		l_texto = " - Ingresaron y No Terminaron "
	case "3":
		l_sql = l_sql & " "
	end select	
	if trim(l_filtro)<>"" then
		l_sql = l_sql & "AND " & l_filtro
	end if	
	l_sql = l_sql & " ORDER BY evacab.cabaprobada, " & l_orden	& " , evadetevldor.evatevnro, evasecc.orden " 
	'response.write(l_sql)
	rsOpen l_rs, cn, l_sql, 0 
	if l_rs.eof then%>
	<tr>
		 <td colspan="9">No hay datos para la selecci&oacute;n realizada.</td>
	</tr>
	<%else
	    l_empleg=0
	    l_evatevnro=0
	    l_evaseccnro=0
	    l_cantidad = 0
	    l_cabaprobada = 9
	    l_listamail = "0"
	    
	    l_nohabilitado=0
	    l_noingreso=0
	    l_notermino=0
		
		do until l_rs.eof
			
			' si quiere for.aprobados y no aprobados poner el titulo
			'if l_obligatoria=2 then
				if l_rs("cabaprobada") <> l_cabaprobada then
					select case l_rs("cabaprobada")
					case -1:%>
						<tr>
						<td colspan="9"><b>Evaluaciones APROBADAS<%=l_texto%></b></td>
						</tr>
					<%
					case 0:%>
						<tr>
						<td colspan="9"><b>Evaluaciones NO APROBADAS<%=l_texto%></b></td>
						</tr>
					<%
					end select
					l_cabaprobada = l_rs("cabaprobada")
				end if	
			'end if
				
			if l_evacabnro <> l_rs("evacabnro") then
				l_cantidad = l_cantidad +1
			end if
			if l_rs("evaoblig")=-1 then
				l_evaoblig="SI"
			else
				l_evaoblig="NO"
			end if
			
			if (isnull(l_rs("fechahab")) or trim(l_rs("fechahab"))="")  AND isnull(l_rs("fechacar")) then
			l_color= "#FFFF00" ' "#dffff"
			l_letra= "<font color=black>"
			else
				if (isnull(l_rs("fechaing")) or trim(l_rs("fechaing"))="") AND isnull(l_rs("fechacar")) then
				'l_color= "#B0E0E6"
				l_color= "#FF0000"
				l_letra= "<font color=white>"
				else
					if isnull(l_rs("fechacar")) or trim(l_rs("fechacar"))="" then
					l_color= "orange" ' "#dffff"
					l_letra= "<font color=black>"
					else
					l_color= "#00FF00"	
					l_letra= "<font color=black>"
					end if
				end if	
			end if	
		
		if l_mostrar =0 then ' DETALLADO%>	
		
		<tr style="background-color:'<%=l_color%>'" onclick="Javascript:Seleccionar(this,<%=l_rs("ternro")%>)">
			<%if int(l_empleg) <> l_rs("empleg") then%>
				<td nowrap><%=l_letra%><%=l_rs("empleg")%></td>
				<td style="width:50"><%=l_letra%><%=l_rs("terape")%>,&nbsp;<%=l_rs("ternom")%></td>
				<%if l_cabaprobada=-1 then
					if (isnull(l_fechaapro) or trim(l_fechaapro)="") then%>
					<td nowrap><%=l_letra%>APROBADA</td>
					<%else%>
					<td nowrap><%=l_letra%><%=l_fechaapro%> </td>
					<%end if%>
				<%else%>
					<td nowrap><%=l_letra%>NO APROBADA</td>
				<%end if%>	
			<%else%>
				<td colspan=3>&nbsp;</td>
			<%end if%>
			
			<%if int(l_evatevnro) <> l_rs("evatevnro") or int(l_empleg) <> l_rs("empleg") then
				if trim(l_rs("evaternro"))<>"" and not isnull(l_rs("evaternro")) then
					l_listamail = l_listamail &"," & l_rs("evaternro")
					%><script>document.datos.listamail.value = '<%=l_listamail%>';</script><%
				end if%>
				<td style="width:100"><%=l_letra%><%=l_rs("evatevdesabr")%>:&nbsp;<%=l_rs("evaterape")%>,&nbsp;<%=l_rs("evaternom")%></td>
			<%else%>
			<td>&nbsp;</td>
			<%end if%>
	
			<%if int(l_evaseccnro) <> l_rs("evaseccnro") then%>
				<td style="width:200"><%=l_letra%><%=l_rs("titulo")%> </td>
				
				<%if l_obligatoria = 2 or l_obligatoria = "" then%>
					<td nowrap align=center><%=l_letra%><%=l_evaoblig%> </td>
				<%end if%>
				
				<%if (isnull(l_rs("fechahab")) or trim(l_rs("fechahab"))="")  And isnull(l_rs("fechacar")) then
					l_nohabilitado=l_nohabilitado + 1
					if trim(leyendas(0)) = trim(l_rs("evatevdesabr")) then
						eval(0)=eval(0) + 1
					else
						auto(0)=auto(0) + 1
					end if
					%>
					<td nowrap align=center><%=l_letra%>NO HABILITADO</td>
					<td nowrap align=center><%=l_letra%>&nbsp;-&nbsp;</td>
					<td nowrap align=center><%=l_letra%>&nbsp;-&nbsp;</td>
				<% else%>
					<td nowrap><%=l_letra%><%=l_rs("fechahab")%> </td>
				<%if isnull(l_rs("fechaing")) or trim(l_rs("fechaing"))="" AND (NOT isnull(l_rs("fechahab")) AND  trim(l_rs("fechahab"))<>"")  then
					l_noingreso     = l_noingreso     + 1
					if trim(leyendas(0)) = trim(l_rs("evatevdesabr")) then
						eval(1)=eval(1) + 1
					else
						auto(1)=auto(1) + 1
					end if
					%>
					<td nowrap align=center><%=l_letra%>NO INGRESO</td>
					<td nowrap align=center><%=l_letra%>&nbsp;-&nbsp;</td>
					<%else%>
					<td nowrap><%=l_letra%><%=l_rs("fechaing")%> </td>
						<%if isnull(l_rs("fechacar")) or trim(l_rs("fechacar"))="" then
						l_notermino=l_notermino+1
						if trim(leyendas(0)) = trim(l_rs("evatevdesabr")) then
							eval(2)=eval(2) + 1
						else
							auto(2)=auto(2) + 1
						end if
						%>
						<td nowrap align=center><%=l_letra%>NO TERMINO</td>
						<%else%>
						<td nowrap><%=l_letra%><%=l_rs("fechacar")%> </td>
						<%end if%>
					<%end if%>
				<%end if%>
				

			<%else%>
				<td colspan=8>&nbsp;</td>
			<%end if%>
		</tr>
		<%
		else ' RESUMEN 
		
			if isnull(l_rs("fechahab")) or trim(l_rs("fechahab"))=""  AND isnull(l_rs("fechacar")) then
				l_nohabilitado=l_nohabilitado+1
					if trim(leyendas(0)) = trim(l_rs("evatevdesabr")) then
						eval(0)=eval(0) + 1
					else
						auto(0)=auto(0) + 1
					end if
			else
				if isnull(l_rs("fechaing")) or trim(l_rs("fechaing"))="" AND (NOT isnull(l_rs("fechahab")) AND  trim(l_rs("fechahab"))<>"") then
					l_noingreso=l_noingreso+1
					if trim(leyendas(0)) = trim(l_rs("evatevdesabr")) then
						eval(1)=eval(1) + 1
					else
						auto(1)=auto(1) + 1
					end if
				else
					if isnull(l_rs("fechacar")) or trim(l_rs("fechacar"))="" then
						l_notermino=l_notermino+1
						if trim(leyendas(0)) = trim(l_rs("evatevdesabr")) then
							eval(2)=eval(2) + 1
						else
							auto(2)=auto(2) + 1
						end if
					end if
				end if
			end if
		end if ' si es detallado o resumen
		
		l_empleg	= l_rs("empleg")
		l_terape	= l_rs("terape")
		l_ternom	= l_rs("ternom")
		l_evaterape = l_rs("evaterape")
		l_evaternom = l_rs("evaternom")
		l_evatevdesabr= l_rs("evatevdesabr")
	    l_evatevnro = l_rs("evatevnro")
	    l_evaseccnro= l_rs("evaseccnro")
		l_evacabnro = l_rs("evacabnro")
		l_ternro    = l_rs("ternro")
		l_fechaapro = l_rs("fechaapro") 
		
				
		l_rs.MoveNext
		if not l_rs.eof then
		
		if l_mostrar<>0 then ' RESUMEN
			'Response.Write ("<script>alert('"&l_evaterape& " -- " & l_rs("evaterape")&"')</script>")
			'Response.Write ("<script>alert('"&l_evaterapeant&"')</script>")
			if int(l_empleg) <> l_rs("empleg") or int(l_evatevnro) <> l_rs("evatevnro") then
			
				if l_noingreso<>0 then
					l_color= "#FF0000"
					l_letra= "<font color=white>"
				else
					if l_notermino<>0 then
						l_color= "orange" ' "#dffff"
						l_letra= "<font color=black>"
					else
						if l_nohabilitado<>0 then
							l_color= "#FFFF00" ' "#dffff"
							l_letra= "<font color=black>"
						end if
					end if
				end if%>
				<tr style="background-color:'<%=l_color%>'" onclick="Javascript:Seleccionar(this,<%=l_rs("ternro")%>)">
				<%if int(l_empleg) <> l_rs("empleg") then%>
					<td nowrap><%=l_letra%><%=l_empleg%></td>
					<td style="width:50"><%=l_letra%><%=l_terape%>,&nbsp;<%=l_ternom%></td>
					<%if l_cabaprobada=-1 then
							if (isnull(l_fechaapro) or trim(l_fechaapro)="") then%>
							<td nowrap><%=l_letra%>APROBADA</td>
							<%else%>
							<td nowrap><%=l_letra%><%=l_fechaapro%> </td>
							<%end if%>
					<%else%>
							<td nowrap><%=l_letra%>NO APROBADA</td>
					<%end if%>	
					<td style="width:100"><%=l_letra%><%=l_evatevdesabr%>:&nbsp;<%=l_evaterape%>,&nbsp;<%=l_evaternom%></td>
					<%if trim(l_rs("evaternro"))<>"" and not isnull(l_rs("evaternro")) then
						l_listamail = l_listamail &"," & l_rs("evaternro")
						%>
						<script>document.datos.listamail.value = '<%=l_listamail%>';</script>
					<%end if%>
				
					<td ALIGN=center nowrap><%=l_letra%><%if l_nohabilitado=0 then%>--<%else%><%=l_nohabilitado%><%end if%></td>
					<td ALIGN=center nowrap><%=l_letra%><%if l_noingreso=0 then%>--<%else%><%=l_noingreso%><%end if%></td>
					<td ALIGN=center nowrap><%=l_letra%><%if l_notermino=0 then%>--<%else%><%=l_notermino%><%end if%></td>					</tr>
				  <%l_nohabilitado=0
					l_noingreso=0
					l_notermino=0
				else ' es el mismo empleado
					if int(l_evatevnro) <> l_rs("evatevnro") then%>
						<td nowrap><%=l_letra%><%=l_empleg%></td>
						<td style="width:50"><%=l_letra%><%=l_terape%>,&nbsp;<%=l_ternom%></td>
						<%if l_cabaprobada=-1 then
							if (isnull(l_fechaapro) or trim(l_fechaapro)="") then%>
							<td nowrap><%=l_letra%>APROBADA</td>
							<%else%>
							<td nowrap><%=l_letra%><%=l_fechaapro%> </td>
							<%end if%>
						<%else%>
							<td nowrap><%=l_letra%>NO APROBADA</td>
						<%end if%>	
						<%if trim(l_rs("evaternro"))<>"" and not isnull(l_rs("evaternro")) then
						l_listamail = l_listamail & "," & l_rs("evaternro")
						%>
						<script>document.datos.listamail.value = '<%=l_listamail%>';</script>
						<%end if%>
						<td style="width:100"><%=l_letra%><%=l_evatevdesabr%>:&nbsp;<%=l_evaterape%>,&nbsp;<%=l_evaternom%></td>
						<td ALIGN=center nowrap><%=l_letra%><%if l_nohabilitado=0 then%>--<%else%><%=l_nohabilitado%><%end if%></td>
						<td ALIGN=center nowrap><%=l_letra%><%if l_noingreso=0 then%>--<%else%><%=l_noingreso%><%end if%></td>
						<td ALIGN=center nowrap><%=l_letra%><%if l_notermino=0 then%>--<%else%><%=l_notermino%><%end if%></td>
					</tr>
					<%l_nohabilitado=0
						l_noingreso=0
						l_notermino=0
					end if ' el evaluador es ditinto
				end if ' es distinto empleado
			end if  ' cambio empleado o evaluador
			end if ' Resumen
			 
			if l_rs("cabaprobada") <> l_cabaprobada then
				select case int(l_cabaprobada)
				case 0:%>
			    <tr>
				<td colspan="10"><b>Cantidad de Evaluaciones No Aprobadas<%=l_texto%>=&nbsp;<%=l_cantidad%></b></td>
				</tr>
				<%
				case -1:%>
			    <tr>
				<td colspan="10"><b>Cantidad de Evaluaciones Aprobadas<%=l_texto%>=&nbsp;<%=l_cantidad%></b></td>
				</tr>
				<%
				end select
				l_cantidad=0
			end if	
		else
			'Response.Write ("<script>alert('"&l_evatevnroant&"')</script>")
			'Response.Write ("<script>alert('"&l_evaterape&"')</script>")
			'Response.Write ("<script>alert('"&l_evaterapeant&"')</script>")
			if l_mostrar<>0 then ' RESUMEN
					if l_noingreso<>0 then
						l_color= "#FF0000"
						l_letra= "<font color=white>"
					else
						if l_notermino<>0 then
							l_color= "orange" ' "#dffff"
							l_letra= "<font color=black>"
						else
							if l_nohabilitado<>0 then
								l_color= "#FFFFCC" ' "#dffff"
								l_letra= "<font color=black>"
							end if
						end if
					end if%>
					<tr style="background-color:'<%=l_color%>'" onclick="Javascript:Seleccionar(this,<%=l_ternro%>)">
						<td nowrap><%=l_letra%><%=l_empleg%></td>
						<td style="width:50"><%=l_letra%><%=l_terape%>,&nbsp;<%=l_ternom%></td>
						<%if l_cabaprobada=-1 then
							if (isnull(l_fechaapro) or trim(l_fechaapro)="") then%>
							<td nowrap><%=l_letra%>APROBADA</td>
							<%else%>
							<td nowrap><%=l_letra%><%=l_fechaapro%> </td>
							<%end if%>
						<%else%>
							<td nowrap><%=l_letra%>NO APROBADA</td>
						<%end if%>	
						<td style="width:100"><%=l_letra%><%=l_evatevdesabr%>:&nbsp;<%=l_evaterape%>,&nbsp;<%=l_evaternom%></td>
						<td ALIGN=center nowrap><%=l_letra%><%if l_nohabilitado=0 then%>--<%else%><%=l_nohabilitado%><%end if%></td>
						<td ALIGN=center nowrap><%=l_letra%><%if l_noingreso=0 then%>--<%else%><%=l_noingreso%><%end if%></td>
						<td ALIGN=center nowrap><%=l_letra%><%if l_notermino=0 then%>--<%else%><%=l_notermino%><%end if%></td>					</tr>
					  <%l_nohabilitado=0
						l_noingreso=0
						l_notermino=0
			end if ' Resumen
		
			select case int(l_cabaprobada)
			case 0:%>
				 <tr>
				<td colspan="10"><b>Cantidad de Evaluaciones No Aprobadas<%=l_texto%>=&nbsp;<b><%=l_cantidad%></b></td>
				</tr>
				<%
			case -1:%>
			    <tr>
				<td colspan="10" ><b>Cantidad de Evaluaciones Aprobadas<%=l_texto%>=&nbsp;<b><%=l_cantidad%></b></td>
				</tr>
				<%
			end select
			l_cantidad=0
		end if
		
		loop
		
	end if ' del if l_rs.eof
	l_rs.Close
	set l_rs = nothing
	cn.Close	
	set cn = nothing
else%>
	<tr>
		 <td colspan="9">Seleccione un Evento</td>
	</tr>
<%end if
		
		
			'Response.Write("<script>alert('"& leyendas(0)&"');</script>")
			'Response.Write("<script>alert('"& eval(0)&"');</script>")
			'Response.Write("<script>alert('"& eval(1)&"');</script>")
			'Response.Write("<script>alert('"& eval(2)&"');</script>")
			'Response.Write("<script>alert('"& leyendas(1)&"');</script>")
			'Response.Write("<script>alert('"& auto(0)&"');</script>")
			'Response.Write("<script>alert('"& auto(1)&"');</script>")
			'Response.Write("<script>alert('"& auto(2)&"');</script>")
		
%>

</table>

<table width="100%" height="300">
  <tr>
	<td height="20" colspan="13" width="100%"></td>
  </tr>
  <tr>
   <td width="5%">
     <br>
   </td>
   <td width="90%">
     <object align="middle" id=ChartSpace1 classid=CLSID:0002E500-0000-0000-C000-000000000046 style="width:100%; height:350; background-color: Aqua;"></object>
   </td>
   <td width="5%">
     <br>
   </td>
  </tr>
  <tr><td colspan="13">&nbsp;</td></tr>
</table>
<script language="vbs">

sub window_onload()
    call generar()
end sub


sub generar()
   
	Dim grafico
    Dim categories(3) ' Valores que van en el eje X
	Dim values(3)    ' Valores que van en el eje Y
	
	set grafico = ChartSpace1

    ' Defino los valores del eje X
    categories(0) =  "No Habilitados"
    categories(1) =  "No Ingresaron"
    categories(2) =  "No Terminaron"
	
    ' Borro el grafico y agrego uno nuevo
    grafico.Clear
    grafico.Charts.Add
    Set c = grafico.Constants	

	' Seteo el tipo de grafico
	grafico.Charts(0).Type = parent.document.all.tipografico.value
			
    ' Agrego tres series al grafico
    grafico.Charts(0).SeriesCollection.Add
    grafico.Charts(0).SeriesCollection.Add
    
	'-------------------------------------------------------
    ' Seteo los valores de la serie 0
    values(0) = <%= eval(0) %>
    values(1) = <%= eval(1) %>
	values(2) = <%= eval(2) %>
	'Seteo algunos parametros de la serie 0
    grafico.Charts(0).SeriesCollection(0).Caption = "<%=leyendas(0)%>"
	grafico.Charts(0).SeriesCollection(0).SetData c.chDimCategories, c.chDataLiteral, categories
    grafico.Charts(0).SeriesCollection(0).SetData c.chDimValues, c.chDataLiteral, values

	'-------------------------------------------------------
    ' Seteo los valores de la serie 1
    values(0) = <%= auto(0)  %>  
    values(1) = <%= auto(1)  %>
	values(2) = <%= auto(2)  %>
	'Seteo algunos parametros de la serie 1
    grafico.Charts(0).SeriesCollection(1).Caption = "<%=leyendas(1)%>"
	grafico.Charts(0).SeriesCollection(1).SetData c.chDimCategories, c.chDataLiteral, categories
    grafico.Charts(0).SeriesCollection(1).SetData c.chDimValues, c.chDataLiteral, values


	' Seteo algunos parametros del grafico
      grafico.Charts(0).HasLegend = true
	' ChartSpace1.Charts(0).Axes(c.chAxisPositionLeft).NumberFormat = "0%"
    ' ChartSpace1.Charts(0).Axes(c.chAxisPositionLeft).MajorUnit = 0.1

end sub

</script>

</body>
</html>
