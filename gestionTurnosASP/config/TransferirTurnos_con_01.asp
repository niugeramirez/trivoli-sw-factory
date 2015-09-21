<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 
'Archivo: contracts_con_01.asp
'Descripción: ABM de Contracts
'Autor : Raul Chinestra
'Fecha: 28/11/2007

Dim l_rs
Dim l_rs2
Dim l_rs3
Dim l_sql
Dim l_sql2
Dim l_sql3
Dim l_filtro
Dim l_cabnro
Dim l_orden
Dim l_sqlfiltro
Dim l_sqlorden
Dim l_totvol
Dim l_cant

Dim l_primero
Dim l_primeravez

Dim l_hd
Dim l_md
Dim l_hh
Dim l_mh  

Dim l_fechadesde
Dim l_fechahasta
Dim l_fechaaux

Dim l_horadesde
Dim l_horahasta

Dim l_DiasSemana

l_fechadesde = request("fechadesde")
l_fechahasta = request("fechahasta")
'response.write l_fechadesde
'response.write l_fechahasta

l_hd = request("hd")
l_md = request("md")
l_hh = request("hh")
l_mh = request("mh")

l_DiasSemana = request("DiasSemana")

l_horadesde = l_hd & ":" & l_md
l_horahasta = l_hh & ":" & l_mh


l_filtro = request("filtro")
l_cabnro = request("cabnro")
l_orden  = request("orden")


Function diasemana(fecha)


	Select Case weekday(fecha)
		Case 1
			diasemana = "#FFFF80"
		Case 2
			diasemana = "#FF0080"
		Case 3
			diasemana = "#FF1180"
		Case 4
			diasemana = "#FF2280"
		Case 5
			diasemana = "#FF3380"
		Case 6
			diasemana = "#FF4480"
		Case 7
			diasemana = "#FF5580"
	End Select
End Function

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Transferir Turnos</title>
</head>

<script>
var jsSelRow = null;

function Deseleccionar(fila){
	fila.className = "MouseOutRow";
}

function Seleccionar(fila,cabnro){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	};
	document.datos.cabnro.value = cabnro;
	fila.className = "SelectedRow";
	jsSelRow = fila;
}


</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="//javascript:parent.Buscar();">
<table >
<% 

l_filtro = replace (l_filtro, "*", "%")


Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")

l_fechaaux = cdate(l_fechadesde)

'response.write "l_fechaaux "&l_fechaaux&"<br/>" &"<br/>" 
'response.write "l_fechahasta "&l_fechahasta&"<br/>" &"<br/>" 

do while cdate(l_fechaaux) <= cdate(l_fechahasta)

  'response.write "l_DiasSemana "&l_DiasSemana&"<br/>" &"<br/>" 
  'response.write "weekday (cdate( l_fechaaux )) "&weekday (cdate( l_fechaaux ))&"<br/>" &"<br/>" 

  if instr(l_DiasSemana, weekday (cdate( l_fechaaux ))) <> 0 then
  %>
    <tr>
        <th align="left" colspan="2" ><%= l_fechaaux %></th>
    </tr> 
  <%
  

  l_sql2 = "SELECT  recursosreservables.descripcion, recursosreservables.id"
  l_sql2 = l_sql2 & " FROM calendarios "
  l_sql2 = l_sql2 & " INNER JOIN recursosreservables ON recursosreservables.id = calendarios.idrecursoreservable   "
  l_sql2 = l_sql2 & " WHERE " & l_filtro & " "
  l_sql2 = l_sql2 & " AND calendarios.id not in (select turnos.idcalendario from turnos)"
  l_sql2 = l_sql2 & " AND calendarios.estado = 'ACTIVO'"
  l_sql2 = l_sql2 & " AND CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  = " & cambiafecha( l_fechaaux ,true,1)
  l_sql2 = l_sql2 & " AND CONVERT(VARCHAR(5), fechahorainicio, 108) >= '" & l_horadesde & "'"   
  l_sql2 = l_sql2 & " AND CONVERT(VARCHAR(5), fechahorainicio, 108) <= '" & l_horahasta & "'"
  l_sql2 = l_sql2 & " GROUP BY recursosreservables.descripcion , recursosreservables.id"
  l_sql2 = l_sql2 & " ORDER BY recursosreservables.descripcion "
  
  'response.write "sql2 "&l_sql2&"<br/>" &"<br/>" 
  
  rsOpen l_rs2, cn, l_sql2, 0
  l_primeravez = true
  do until l_rs2.eof
	if l_primeravez = true then
	%>
    <tr>
        <th>Medico</th>
        <th>Turnos Disponibles</th>
    </tr>	
	<%
		l_primeravez = false
	end if 
	%>	
    <tr>
        <td nowrap align="left" width="15%" ><%= l_rs2("descripcion") %></td>        
    
	<%
	
	  l_sql3 = "SELECT  calendarios.id,   CONVERT(VARCHAR(5), fechahorainicio, 108) AS horainicio , fechahorainicio , CONVERT(VARCHAR(10), fechahorainicio, 101) AS Fecha , recursosreservables.descripcion"
	  l_sql3 = l_sql3 & " FROM calendarios "
	  l_sql3 = l_sql3 & " INNER JOIN recursosreservables ON recursosreservables.id = calendarios.idrecursoreservable   "
	  l_sql3 = l_sql3 & " WHERE " & l_filtro & " "
	  l_sql3 = l_sql3 & " AND calendarios.idrecursoreservable = " & l_rs2("id")
	  l_sql3 = l_sql3 & " AND calendarios.id not in (select turnos.idcalendario from turnos)"
	  l_sql3 = l_sql3 & " AND calendarios.estado = 'ACTIVO'"
      l_sql3 = l_sql3 & " AND CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101)  = " & cambiafecha( l_fechaaux ,true,1)	  
	  l_sql3 = l_sql3 & " ORDER BY recursosreservables.descripcion "
      'response.write "sql3 "& l_sql3	&"<br/>" &"<br/>"  
	  
	  rsOpen l_rs3, cn, l_sql3, 0
	  
	  if not l_rs3.eof then
	  %>
	   <td nowrap align="left"  >
	  <%  
	  end if
	  
	  do until l_rs3.eof
			'l_cant = l_cant + 1
			'response.write   "rrrrr"
			if l_rs3("horainicio") >= l_horadesde  and l_rs3("horainicio") <= l_horahasta and instr(l_DiasSemana, weekday (cdate( l_rs3("fechahorainicio"))) ) <> 0 then
			%>	
	    
	        	<a href="Javascript:parent.abrirVentana('TransferirTurnos_con_02.asp?Tipo=A&ant=<%= l_cabnro %>&nuevo=<%= l_rs3("id")%>' ,'',600,300);"><%'= cambiafecha(l_rs3("Fecha"),"","")%><%= l_rs3("horainicio")%>&nbsp;<%'= weekday(cdate(l_rs2("Fecha"))) %></a>
	    
			<%
			end if
			l_rs3.MoveNext
	  loop	
	 
	   l_rs3.Close	
	
	  %>
	  </td>
	  </tr>
	  <%
	
	
		l_rs2.MoveNext
  loop	
 
   l_rs2.Close
  
  
  end if

  l_fechaaux = cdate(l_fechaaux) + 1
 	
loop


l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
<script>    
	parent.parent.ActPasos(<%= l_primero %>,"","MENU");
    parent.parent.datos.pasonro.value = <%= l_primero %>;
</script>

</table>
<form name="datos" method="post">
<input type="hidden" name="cabnro" value="<%= l_cabnro %>">
<input type="hidden" name="orden" value="<%= l_orden %>">
<input type="hidden" name="filtro" value="<%= l_filtro %>">
</form>
</body>
</html>
