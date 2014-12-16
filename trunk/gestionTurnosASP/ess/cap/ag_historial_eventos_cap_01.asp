<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: rep_historial_eventos_cap_01.asp
Autor: Gustavo Ring
Creacion: 28/05/2007
Descripcion: Muestra el historial de eventos del empleado
Modificacion:
-----------------------------------------------------------------------------
-->
<% 
on error goto 0


Const l_Max_Lineas_X_Pag = 50
Const l_cantcols = 4

Dim l_rs
Dim l_rs2
Dim l_rs3
Dim l_rs4

Dim l_sql
Dim l_sql2
Dim l_sql3
Dim l_sql4
Dim l_filtro

Dim l_nrolinea
Dim l_nropagina
Dim l_cantreg
Dim i

Dim l_fecha
Dim l_fecha2

Dim l_encabezado
Dim l_corte 
Dim l_corteempleado 

Dim l_tipodia
Dim l_ternro
Dim l_certificado

Dim l_fecdesde 
Dim l_fechasta 

Dim l_tenro1
Dim l_estrnro1
Dim l_tenro2
Dim l_estrnro2
Dim l_tenro3
Dim l_estrnro3

Dim l_eventos

Dim l_detallado

Dim l_orden

Dim l_estr1ant
Dim l_estr2ant
Dim l_estr3ant

Dim l_fecha_sql

Dim l_int_esp 
	l_int_esp= false
Dim l_int_corte
	l_int_corte= false
	
dim l_total 
dim l_totalgral
    l_totalgral = 0

Dim l_totalgrupo(3)
Dim l_totalgrupohs(3)

Dim l_hay_datos

Dim l_cantidademp

Dim l_filtro_join

Dim mostre

Dim l_subtotal_horas
Dim l_totalgral_horas

l_subtotal_horas = 0
l_totalgral_horas = 0

l_ternro = request.QueryString("ternro")

' Imprime el encabezado de cada pagina
sub encabezado(titulo)
%>
	<table>
	<tr>
		<td align="center" colspan="<%= l_cantcols%>">
		<table>
			<tr>
		       	<td align="right" width="10%"> 
					P&aacute;gina: <%= l_nropagina%>
				</td>				
			</tr>
		</table>
		</td>				
	</tr>
</u>
<%
end sub 'encabezado

' Carga los datos en RS2
sub cargar_datos

	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

	l_sql = "SELECT cap_evento.evenro, evecodext, evedesabr, curdesabr, cap_evento.evereqasi, cap_evento.eveporasi "
	l_sql = l_sql & " FROM cap_candidato "
    l_sql = l_sql & " INNER JOIN cap_evento ON cap_evento.evenro = cap_candidato.evenro "
	l_sql = l_sql & " INNER JOIN cap_curso ON cap_curso.curnro = cap_evento.curnro "	
	
	l_sql = l_sql & " INNER JOIN cap_eventomodulo ON cap_evento.evenro= cap_eventomodulo.evenro "
	l_sql = l_sql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro "
    
	l_sql = l_sql & " WHERE cap_candidato.ternro =" & l_ternro & " AND cap_candidato.conf = -1"
	
	if l_fecdesde <> "" then
		l_sql = l_sql & " AND cap_calendario.calfecha >= " & cambiafecha(l_fecdesde,"","")
	end if
	if l_fechasta <> "" then
		l_sql = l_sql & " AND cap_calendario.calfecha <= " & cambiafecha(l_fechasta,"","")
	end if
	l_sql = l_sql & " GROUP BY cap_evento.evenro, evecodext, evedesabr, curdesabr, cap_evento.evereqasi, cap_evento.eveporasi "
    
	l_sql = l_sql & " ORDER BY evecodext"

	'response.write l_sql & "<br>"
	'response.end

	rsOpen l_rs2, cn, l_sql, 0 
	
	if (not l_rs2.eof) and (l_estr1ant="vacio") then
		l_estr1ant= "lleno"
		if l_tenro1 <> "" and l_tenro1 <> "0" then
			l_estr1ant = l_rs("estrnro1")
		end if
		if l_tenro2 <> "" and l_tenro2 <> "0" then
			l_estr2ant = l_rs("estrnro2")
		end if
		if l_tenro3 <> "" and l_tenro3 <> "0" then
			l_estr3ant = l_rs("estrnro3")
		end if
	end if	

end sub 'cargar_datos

'Imprime el titulo de la seccion de cada empleado
sub titulo_empleado(Empleado,apellido,nombre)
	l_nrolinea = l_nrolinea+1
	%>				
	<tr>
       <td STYLE="border : thick solid 1;" align="left" colspan="<%= l_cantcols%>">
		 <B><%= Empleado & " - " & apellido & ", " & nombre%></B>
  	   </td>				
    </tr>		    
	<%
end sub 'titulo_empleado

'Calcula el total de horas del evento evenro
function horas_del_evento(evenro)
dim f_rs
dim f_sql
dim l_dummy
dim l_totalhs
	l_totalhs = Cint(0)
	l_dummy = Cint(0)
	f_sql = " SELECT cap_eventomodulo.evenro, calfecha, calhordes, calhorhas "
	f_sql = f_sql & " FROM cap_eventomodulo "
	f_sql = f_sql & " INNER JOIN cap_calendario ON cap_calendario.evmonro = cap_eventomodulo.evmonro "
	f_sql = f_sql & " WHERE cap_eventomodulo.evenro = " & evenro
	Set f_rs = Server.CreateObject("ADODB.RecordSet")
	rsOpen f_rs, cn, f_sql, 0 
	do until f_rs.eof
		l_dummy = datediff("n",cdate(mid(f_rs("calhordes"),1,2)&":"& mid(f_rs("calhordes"),3,2)),cdate(mid(f_rs("calhorhas"),1,2)&":"& mid(f_rs("calhorhas"),3,2)))
		l_totalhs = l_totalhs + replace(Round((l_dummy/60),2),",",".")
		f_rs.MoveNext
	loop
	f_rs.close
	set f_rs = nothing
	horas_del_evento = l_totalhs
end function

'Imprime uno de los registros de cada empleado
sub mostrar_datos

Dim l_desc_tipo
Dim l_desc_tipo_abr
Dim l_desc_cert
Dim l_horas

	l_desc_tipo = ""
	l_desc_tipo_abr = ""
	
	 if l_eventos = 0 then 
	%>
		<tr>
			<td align="left" nowrap><%=l_rs2("evecodext")%></td>
			<td align="left" nowrap><%=l_rs2("evedesabr")%></td>
			<td align="left" nowrap><%=l_rs2("curdesabr")%></td><%
			l_horas = horas_del_evento(l_rs2("evenro"))
			l_subtotal_horas = l_subtotal_horas + l_horas%>
			<td align="right" nowrap><%= l_horas%></td>
			<% l_total = 	l_total + 1  
			   mostre = true
			%>	
		</tr><% 
	else
		asistencia_al_evento	
	end if
end sub 'mostrar datos

' ------------------------------------------------------------------------------------
' funcion para imprimir los totales de cada estructura
 function totalesgrupal

 dim l_nrolineaant
 l_nrolineaant = l_nrolinea
if  l_tenro3 <> "" and l_tenro3 <> "0" then
	if l_estr3ant <> l_rs("estrnro3") or l_estr2ant <> l_rs("estrnro2") or l_estr1ant <> l_rs("estrnro1") then
		l_nrolinea = l_nrolinea+2
		'Total de horas
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>----Total Horas "& desc_estructura(l_tenro3, l_estr3ant) &"<b></td>"
		response.write "<td align=right ><b>" & l_totalgrupohs(3) & "----</b></td>"
		response.write "</tr>"
		'Total de eventos
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>----Total Eventos "& desc_estructura(l_tenro3, l_estr3ant) &"<b></td>"
		if l_totalgrupo(3) = 0 then
			response.write "<td align=center>  </td>"		
		else
			response.write "<td align=right ><b>" & l_totalgrupo(3) & "----</b></td>"
		end if
		response.write "</tr>"

		l_totalgrupohs(3) = 0
		l_totalgrupo(3) = 0
	end if
 end if
if  l_tenro2 <> "" and l_tenro2 <> "0" then
	if l_estr2ant <> l_rs("estrnro2") or l_estr1ant <> l_rs("estrnro1") then
		l_nrolinea = l_nrolinea+2
		'Total de horas
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>--Total Horas "& desc_estructura(l_tenro2, l_estr2ant) &"<b></td>"
		response.write "<td align=right><b>" & l_totalgrupohs(2) & "--</b></td>"
		response.write "</tr>"
		'total de eventos
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>--Total Eventos "& desc_estructura(l_tenro2, l_estr2ant) &"<b></td>"
		if l_totalgrupo(2) = 0 then
			response.write "<td align=center>  </td>"		
		else
			response.write "<td align=right><b>" & l_totalgrupo(2) & "--</b></td>"
		end if
		response.write "</tr>"

		l_totalgrupohs(2) = 0
		l_totalgrupo(2) = 0
	end if
 end if
if  l_tenro1 <> "" and l_tenro1 <> "0" then
 	if l_estr1ant <> l_rs("estrnro1") then
		l_nrolinea = l_nrolinea+2
		'total horas
		response.write "<tr><td  align=left colspan=" & l_cantcols - 1 & "><b>Total Horas "& desc_estructura(l_tenro1, l_estr1ant) &"<b></td>"
		response.write "<td align=right ><b>" & l_totalgrupohs(1) & "</b></td>"
		response.write "</tr>"
		'total eventos
		response.write "<tr><td  align=left colspan=" & l_cantcols - 1 & "><b>Total Eventos "& desc_estructura(l_tenro1, l_estr1ant) &"<b></td>"
		if l_totalgrupo(1) = 0 then
			response.write "<td align=center>  </td>"		
		else
			response.write "<td align=right ><b>" & l_totalgrupo(1) & "</b></td>"
		end if
		response.write "</tr>"

		l_totalgrupohs(1) = 0
		l_totalgrupo(1) = 0
	end if
 end if
 if l_nrolinea <> l_nrolineaant then ' es que puso algun total
	%>				
	<tr>
		<td colspan="<%=l_cantcols%>">&nbsp;</td>
	</tr>
<%
 end if
 end function

'----------------------------------------------------------------------------------
'  grupales cuando termina el bucle --------
 function totalesgrupalfinal
if  l_tenro3 <> "" and l_tenro3 <> "0" then
		l_nrolinea = l_nrolinea+2
		'total hs
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>----Total Horas "& desc_estructura(l_tenro3, l_estr3ant) &"<b></td>"
	    response.write "<td align=right colspan=1><b>" & l_totalgrupohs(3) & "----</b></td><td></td>"
		response.write "</tr>"
		'total eventos
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>----Total Eventos "& desc_estructura(l_tenro3, l_estr3ant) &"<b></td>"
		if l_totalgrupo(3) = 0 then
			response.write "<td align=center>  </td>"		
		else
		    response.write "<td align=right colspan=1><b>" & l_totalgrupo(3) & "----</b></td><td></td>"
		end if
		response.write "</tr>"

		l_totalgrupohs(3) = 0
		l_totalgrupo(3) = 0
 end if
if  l_tenro2 <> "" and l_tenro2 <> "0" then
		l_nrolinea = l_nrolinea+2
		'total hs
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>--Total Horas "& desc_estructura(l_tenro2, l_estr2ant) &"<b></td>"
		response.write "<td align=right colspan=1><b>" & l_totalgrupohs(2) & "--</b></td><td></td>"
		response.write "</tr>"
		'total eventos
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>--Total Eventos "& desc_estructura(l_tenro2, l_estr2ant) &"<b></td>"
		if l_totalgrupo(2) = 0 then
			response.write "<td align=center>  </td>"		
		else
			response.write "<td align=right colspan=1><b>" & l_totalgrupo(2) & "--</b></td><td></td>"
		end if
		response.write "</tr>"

		l_totalgrupohs(2) = 0
		l_totalgrupo(2) = 0
 end if
if  l_tenro1 <> "" and l_tenro1 <> "0" then
		l_nrolinea = l_nrolinea+2
		'total hs
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>Total Horas "& desc_estructura(l_tenro1, l_estr1ant) &"<b></td>"
		response.write "<td align=right colspan=1><b>" & l_totalgrupohs(1) & "</b></td><td></td>"
		response.write "</tr>"
		'total eventos
		response.write "<tr><td align=left colspan=" & l_cantcols - 1 & "><b>Total Eventos "& desc_estructura(l_tenro1, l_estr1ant) &"<b></td>"
		if l_totalgrupo(1) = 0 then
			response.write "<td align=center>  </td>"		
		else
			response.write "<td align=right colspan=1><b>" & l_totalgrupo(1) & "</b></td><td></td>"
		end if
		response.write "</tr>"

		l_totalgrupohs(1) = 0
		l_totalgrupo(1) = 0
 end if
 end function
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
' Determina si el empledado aprobo la asistencia al evento --------
 sub asistencia_al_evento
    
	Dim l_tot
	Dim l_can
	Dim l_por
    Dim l_portot
	
	' Calculo el numero total de minutos que Debe ir el Participante  
	
	Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
	
	l_sql3 = " SELECT cap_calendario.calnro, calhordes, calhorhas "
	l_sql3 = l_sql3 & " FROM cap_eventomodulo "
	l_sql3 = l_sql3 & " INNER JOIN cap_calendario ON cap_calendario.evmonro = cap_eventomodulo.evmonro "
	l_sql3 = l_sql3 & " INNER JOIN cap_partcal ON cap_partcal.calnro = cap_calendario.calnro AND cap_partcal.ternro = " & l_rs("ternro")
	l_sql3 = l_sql3 & " WHERE cap_eventomodulo.evenro = " & l_rs2("evenro") 
	rsOpen l_rs3, cn, l_sql3, 0 
	l_tot = 0
	l_can = 0
	l_portot = l_rs2("eveporasi")

	do until l_rs3.eof
		
		l_tot = l_tot + datediff("n",cdate(mid(l_rs3("calhordes"),1,2)&":"& mid(l_rs3("calhordes"),3,2)),cdate(mid(l_rs3("calhorhas"),1,2)&":"& mid(l_rs3("calhorhas"),3,2)))
        
		' Calculo el numero total de minutos que Asistio el Empleado  
		
		Set l_rs4 = Server.CreateObject("ADODB.RecordSet")
		
		l_sql4 = " SELECT asipre "
		l_sql4 = l_sql4 & " FROM cap_asistencia "
		l_sql4 = l_sql4 & " WHERE cap_asistencia.ternro = " & l_rs("ternro") & " AND cap_asistencia.calnro = " & l_rs3("calnro")
		
		rsOpen l_rs4, cn, l_sql4, 0 
		
		if not(l_rs4.eof) then
			if  l_rs4("asipre") = -1 then 
				l_can = l_can + datediff("n",cdate(mid(l_rs3("calhordes"),1,2)&":"& mid(l_rs3("calhordes"),3,2)),cdate(mid(l_rs3("calhorhas"),1,2)&":"& mid(l_rs3("calhorhas"),3,2)))							
			end if
		end if
		l_rs4.close
		l_rs3.MoveNext
		
	loop
	l_rs3.Close
	
	if l_tot = 0 then 
	   l_por = 0
	   else l_por = l_can * 100 / l_tot
	end if   
    if l_por >= l_portot then
			if l_eventos = 1  then 
		%>
	                <tr>
					    <% if l_corteempleado then 
		                      titulo_empleado l_rs("empleg"),l_rs("terape") & " " & l_rs("terape2") ,l_rs("ternom") & " " & l_rs("ternom2")
		                  end if	  
		                %>
	                    <td width="20%" align="left"><%= l_rs2("evecodext")%></td>
	                    <td width="40%" align="left"><%= l_rs2("evedesabr")%></td>
	                    <td width="40%" align="left"><%= l_rs2("curdesabr")%></td>
	                </tr>
				   <% l_total = 	l_total + 1  
				      mostre = true
			else mostre = false		  
				   %>	
		<%	end if
		    else if l_eventos = 2  then 
		%>
		            <tr>
   					    <% if l_corteempleado then 
		                      titulo_empleado l_rs("empleg"),l_rs("terape") & " " & l_rs("terape2") ,l_rs("ternom") & " " & l_rs("ternom2")
		                   end if	  
		                %>
	                    <td width="20%" align="left"><%= l_rs2("evecodext")%></td>
	                    <td width="40%" align="left"><%= l_rs2("evedesabr")%></td>
	                    <td width="40%" align="left"><%= l_rs2("curdesabr")%></td>
	                </tr>
					<% l_total = 	l_total + 1 
					   mostre = true 
				    else mostre = false	 
					   %>	
				<% end if %>
	<%	end if 
	
 end sub
'----------------------------------------------------------------------------------
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="../<%= c_EstiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

</script>


<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%


Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_corteempleado = true

l_estr1ant= "vacio"

l_sql = " SELECT * FROM v_empleado WHERE ternro=" & l_ternro
rsOpen l_rs, cn, l_sql, 0 

l_hay_datos = false

    cargar_datos
	'response.end
	
	do while not l_rs2.eof
	
		if l_encabezado then 
			if l_corte then
                response.write "</table><p style='page-break-before:always'></p>"
				l_nrolinea = 1
			end if 		
			
			encabezado("Actividades de Capacitación realizadas por Empleados")
			
			%>
		    <tr>
	    	    <th nowrap>C&oacute;digo</th>
	            <th nowrap>Descripci&oacute;n</th>
			    <th nowrap>Curso</th>

			    <th nowrap>Horas</th>

			</tr>
		<%
			
			l_nrolinea = l_nrolinea+2
		end if	

		if l_corteempleado then 
		    'titulo_empleado l_rs("empleg"),l_rs("terape"),l_rs("ternom")	
			l_hay_datos = true			
		end if
		
		mostrar_datos
		
		l_corteempleado = false	
	
		l_rs2.MoveNext			

		l_nrolinea = l_nrolinea + 1	
		
		if l_nrolinea > l_Max_Lineas_X_Pag then 
			l_corte = true		
			l_encabezado = true
			l_corteempleado = true				
			l_nropagina	= l_nropagina + 1
		else 
			l_encabezado = false
		end if
	loop
	l_corteempleado = true	
	l_rs2.Close
    l_rs.MoveNext
	
	if l_hay_datos AND l_total <> 0 then
		response.write "<tr><td align=""left""><b>Total Horas</b></td><td colspan=" & l_cantcols -1 & " align=""right""><b>" & l_subtotal_horas & "</b></td></tr>"
		response.write "<tr><td align=""left""><b>Total Eventos</b></td><td colspan=" & l_cantcols -1 & " align=""right""><b>" & l_total & "</b></td></tr>"
		l_totalgral_horas = l_totalgral_horas + l_subtotal_horas
		    if l_tenro1 <> "" and l_tenro1 <> "0" then
			   l_totalgrupohs(1) = l_totalgrupohs(1) + l_subtotal_horas
			end if	
			if l_tenro2 <> "" and l_tenro2 <> "0" then
			   l_totalgrupohs(2) = l_totalgrupohs(2) + l_subtotal_horas
			end if	
			if l_tenro3 <> "" and l_tenro3 <> "0" then
			   l_totalgrupohs(3) = l_totalgrupohs(3) + l_subtotal_horas
			end if	
		l_subtotal_horas = 0
		l_nrolinea = l_nrolinea + 2
	end if
	
	l_totalgral = l_totalgral + l_total
	l_total = 0	

if not l_hay_datos then
%>
 <table>
 	<tr>
		<td colspan="<%= l_cantcols %>"><b>No hay datos para la selecci&oacute;n actual.</b>
		</td>
	</tr>
<%
else

'totalesgrupalfinal

%>
 <tr><td colspan="<%= l_cantcols%>"><hr width="100%" size="2"></td></tr>
<%
end if

l_rs.Close
set l_rs = Nothing
set l_rs2 = Nothing
set l_rs3 = Nothing
cn.Close
set cn = Nothing
%>
</table>
</body>
</html>

