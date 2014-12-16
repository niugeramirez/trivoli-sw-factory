<% 
'--------------------------------------------------------------------------
'Archivo       : resumen_datosadm_eva_00.inc
'Descripcion   : resumen de los datos administrativos
'Creacion      : 29 - 12- 2004
'Autor         : Leticia Amadio.
'Modificacion  : 
'--------------------------------------------------------------------------


function resumen_datosadm

	' variables que vienen del reporte que llama a la funcion ---
		'Dim l_evacabnro, l_evaseccnro 
		'l_evacabnro = 272
		'l_evaseccnro = 133
		'Dim l_linea
		'Dim l_sql ' ver si ya vien definida


Dim l_gerente
Dim l_revisor
Dim l_rsproy, l_rsproy1
Set l_rsproy = Server.CreateObject("ADODB.RecordSet")
Set l_rsproy1 = Server.CreateObject("ADODB.RecordSet")

 Dim l_rs1
   Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = " SELECT horas, fechareunion, basereunion, evadatosadm.evldrnro "
	l_sql = l_sql & " FROM evadatosadm "
	l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evadatosadm.evldrnro "
	l_sql = l_sql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
	l_sql = l_sql & " INNER JOIN evadet ON evadet.evacabnro = evacab.evacabnro "
	l_sql = l_sql & " WHERE evadet.evacabnro= " & l_evacabnro
	l_sql = l_sql & " AND evadet.evaseccnro = " & l_evaseccnro
	
	rsOpen l_rs1, cn, l_sql, 0
%>

<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FAF0E6">
<tr>
	<td colspan="8" align="center">	<b> Datos administrativos </b></td>
</tr>
<%
	if not l_rs1.eof then
	    ' buscar datos de proyecto relacionado con evldrnro
		l_sql = "SELECT  estrdabr,  proygerente, proyrevisor, evaproyfdd,  evaclinom, evadetevldor.evatevnro   " 
		l_sql = l_sql & " FROM evaproyecto "
		l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = evaproyecto.estrnro  "  
		   	' nombre del cliente 
		l_sql = l_sql & " INNER JOIN evaengage  ON evaengage.evaengnro = evaproyecto.evaengnro "
		l_sql = l_sql & " INNER JOIN evacliente ON evacliente.evaclinro = evaengage.evaclinro "
		  	' 
		l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
		l_sql = l_sql & " INNER JOIN evacab ON evacab.evaevenro = evaevento.evaevenro "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
   		l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_rs1("evldrnro")
   		
		rsOpen l_rsproy, cn, l_sql, 0
		
		' Selecciona el nombre del gerente y del revisor.
		if not l_rsproy.eof then
			l_sql= " SELECT terape, terape2,ternom, ternom2 "
			l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rsproy("proygerente")
		  	rsOpen l_rsproy1, cn, l_sql, 0
			l_gerente = l_rsproy1("terape") & " " &  l_rsproy1("terape2") & " " & l_rsproy1("ternom") &  " "  & l_rsproy1("ternom2")
			l_rsproy1.Close
			
			l_sql= " SELECT terape, terape2,ternom, ternom2 "
			l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rsproy("proyrevisor")
		  	rsOpen l_rsproy1, cn, l_sql, 0
			l_revisor = l_rsproy1("terape") & " " &  l_rsproy1("terape2") & " " & l_rsproy1("ternom") &  " "  & l_rsproy1("ternom2")
			l_rsproy1.Close
			
		end if
	%>

		<tr>
			<td align="center"><b>Linea de Servicio </b></td>
			<td align="center"><b>Gerente </b></td>
			<td align="center"><b>Revisor </b></td>
			<td align="center"><b>Per&iacute;odo de Revisi&oacute;n </b></td>
			<td align="center"><b>Cliente </b></td>	
			<td align="center"><b>Horas Imputadas</b></td>
			<td align="center"><b>Reuni&oacute;n</b></td>
			<td align="center"><b>Bases para la revisi&oacute;n</b></td>
		</tr>
   		<tr>
		 	<td align="left"><%= l_rsproy("estrdabr")%>&nbsp;</td> 
			<td align="left"><%=l_gerente %></td> 
			<td align="left"> <%= l_revisor%></td>
			<td align="left">Desde el <%=l_rsproy("evaproyfdd")%></td>	
			<td align="left"> <%= l_rsproy("evaclinom")%>&nbsp;</td> 	
			<td valign=top align="center"><%=l_rs1("horas")%></td>
			<td valign=top align="center"><%=l_rs1("fechareunion")%></td>
			<%
			select case l_rs1("basereunion")
			case 1: %>	
				<td valign=top align="center">Moderada</td>
			<% case 2: %>
				<td valign=top align="center">Urgente</td>
			<% end select%>	
		</tr>		

<%
		l_linea = l_linea + 3
		
	else
%>
		<tr>
			<td valign="top" align="center" colspan="8">--&nbsp;</td>
		</tr>
<%
		l_linea = l_linea + 2
	end if
%>
   	
 </table>

<%
  l_rs1.Close
end function
%>
