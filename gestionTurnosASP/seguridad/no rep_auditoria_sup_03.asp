<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--
'-----------------------------------------------------------------------------------
Archivo	    : rep_libro_ley_liq_03.asp
Descripción : Controla el estado del proceso
Autor		: Scarpa D. 
Fecha		: 23/03/2004
Modificado	: 
	15-07-04 - Leticia Amadio - Se cambio en algunas partes la funcion CInt por CLng
------------------------------------------------------------------------------------
-->
<% 
on error goto 0

'Variables base de datos
 Dim l_rs
 Dim l_cm
 Dim l_sql

'Variables uso local
 Dim l_porc
 Dim l_bpronro
 
 Dim l_tiempo_actual
 Dim l_tiempo_restante
 Dim l_tiempo_total
 Dim l_total_emp
 Dim l_restantes_emp
 
 Dim l_desde
 Dim l_hasta
 Dim l_empresa
 Dim l_empnom 
 Dim l_incOperBen
 
 l_bpronro = request("bpronro")
 'l_total_emp = request("totalemp")
 
 'l_bpronro = 3743
 
 'if l_total_emp = "" then
 '   l_total_emp = 0
 'end if
%>
<script src="/rhprox2/shared/js/fn_windows.js"></script>
<html>
<head>
<title>Estado Proceso</title>
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%

function convAHora(milisec)

  Dim seg
  Dim min
  Dim hor
  
  seg = milisec \ 1000
  min = seg \ 60
  hor = min \ 60
  seg = seg mod 60
  min = min mod 60
  
  if seg < 10 then
     seg = "0" & seg
  end if
  
  if min < 10 then
     min = "0" & min
  end if
  
  if hor < 10 then
     hor = "0" & hor
  end if
  
  convAHora = hor & ":" & min & ":" & seg
end function 'convAHora(milisec)

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Busco el estado del proceso
l_sql = "SELECT * FROM batch_proceso WHERE bpronro=" & l_bpronro

l_restantes_emp = 0

rsOpen l_rs, cn, l_sql, 0
if not l_rs.EOF then
   if CStr(l_rs("bprcestado") <> "Procesado") then

	   if isNull(l_rs("bprcprogreso")) then
	      l_porc = "0"   
		  l_tiempo_actual = "0"
	      'l_restantes_emp = 0
	   else
	      if l_rs("bprcprogreso") = "" then
		     l_porc = "0"
			 l_tiempo_actual = "0"
	  	     'l_restantes_emp = 0
		  else
		     l_porc = CLng(l_rs("bprcprogreso"))
			 if l_rs("bprcprogreso") <> "" then
			    l_tiempo_actual = CLng(l_rs("bprctiempo"))
				if isnull(l_rs("bprcempleados")) then
				   'l_restantes_emp = 0
				else
				   if l_rs("bprcempleados") <> "" then
				      'l_restantes_emp = l_rs("bprcempleados")
				   else
				      'l_restantes_emp = 0 
				   end if
				end if
			 else
			    l_tiempo_actual = 0
	            'l_restantes_emp = 0
			 end if
		  end if
	   end if	  
	else
   	   l_porc = "100"
	end if
else
   	l_porc = "100"
end if

l_rs.close

if l_porc <> "0" then
   l_tiempo_total = (CLng(l_tiempo_actual) * 100) / CLng(l_porc)
   l_tiempo_restante = l_tiempo_total - l_tiempo_actual
else
   l_tiempo_total = 0
   l_tiempo_restante = 0
end if

'Me fijo cuantos empleados quedan sin liquidar

'if (l_total_emp = 0) AND (l_restantes_emp <> 0) then
'   l_total_emp = l_restantes_emp
'end if

%>

<script>
function refrescar(){
  window.location = 'rep_auditoria_sup_03.asp?bpronro=<%= l_bpronro%>&desde=<%= l_desde%>&hasta=<%= l_hasta%>&empresa=<%= l_empresa%>&incOperBen=<%= l_incOperBen%>';
}
</script>

<%
l_sql =         "<table align=""center"" width=""100%"" height=""10"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
l_sql = l_sql & "<tr>"
l_sql = l_sql & "<td width=""5"">		  "
l_sql = l_sql & "	  </td>"
if l_porc <> "0" then
   l_sql = l_sql & "	  <td style=""background-color: DarkBlue;"" width=""" & l_porc & "%"" >"
   l_sql = l_sql & "	  </td>				  "
end if
if l_porc <> "100" then
   l_sql = l_sql & "	  <td style=""background-color: Silver"" width=""" & CStr(100 - l_porc) & "%""> "
   l_sql = l_sql & "	  </td>"
end if
l_sql = l_sql & "	  <td width=""5""> "
l_sql = l_sql & "	  </td>"		  
l_sql = l_sql & "	</tr>"	  
l_sql = l_sql & "  </table>"
l_sql = l_sql & " " & l_porc & "%"
%>

<%if CLng(l_porc) = 100 then%>
<script>
 parent.cambiar('<%= l_porc%>','<%= convAHora(l_tiempo_actual)%>','<%= convAHora(l_tiempo_restante) %>','<%= trim(l_sql) %>');
 parent.parent.actualizar('<%= l_bpronro%>');
</script>
<%else%>
<script> 
 parent.cambiar('<%= l_porc%>','<%= convAHora(l_tiempo_actual)%>','<%= convAHora(l_tiempo_restante) %>','<%= trim(l_sql) %>');
 setTimeout("refrescar()", 3000);
</script>
<%end if%>

</body>
</html>





