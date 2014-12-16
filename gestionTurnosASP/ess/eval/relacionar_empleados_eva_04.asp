<% Option Explicit %>

<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<!--
Archivo	    : relacionar_empleados_eva_04.asp
Descripción : Controla el estado del proceso
Autor		: CCCRossi. 
Fecha		: 18-01-2005
Modificado	: 
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
 
 l_bpronro = request("bpronro")
 l_total_emp = request("totalemp")
 
 if l_total_emp = "" then
    l_total_emp = 0
 end if
%>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
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



l_restantes_emp = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")
'Busco el total de empleados
l_sql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro=" & l_bpronro
rsOpen l_rs, cn, l_sql, 0

if l_total_emp = "" then
   l_total_emp = l_rs("total")
else
   l_restantes_emp = l_rs("total")
end if

l_rs.close

'Busco el estado del proceso
l_sql = "SELECT * FROM batch_proceso WHERE bpronro=" & l_bpronro



rsOpen l_rs, cn, l_sql, 0
if not l_rs.EOF then

   if (Ucase(CStr(l_rs("bprcestado"))) = UCase("Procesando")) OR (Ucase(CStr(l_rs("bprcestado"))) = UCase("Pendiente"))then

	   if isNull(l_rs("bprcprogreso")) then
	      l_porc = "0"   
		  l_tiempo_actual = "0"
	   else
	      if l_rs("bprcprogreso") = "" then
		     l_porc = "0"
			 l_tiempo_actual = "0"
		  else
		     l_porc = CLng(l_rs("bprcprogreso"))
			 if l_rs("bprcprogreso") <> "" then
			    l_tiempo_actual = CLng(l_rs("bprctiempo"))
				if isnull(l_rs("bprcempleados")) then
				   l_restantes_emp = 0
				else
				   if l_rs("bprcempleados") <> "" then
				      l_restantes_emp = l_rs("bprcempleados")
				   else
				      l_restantes_emp = 0 
				   end if
				end if
			 else
			    l_tiempo_actual = 0
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

%>

<script>
function refrescar(){
  window.location = 'relacionar_empleados_eva_04.asp?bpronro=<%= l_bpronro%>&totalemp=<%= trim(l_total_emp) %>';
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
 parent.cambiar('<%= l_porc%>','<%= trim(l_total_emp) %>','<%= trim(l_restantes_emp) %>','<%= convAHora(l_tiempo_actual)%>','<%= convAHora(l_tiempo_restante) %>','<%= trim(l_sql) %>');
 alert('Operación Realizada');
 parent.close(); 
 //parent.actualizarEmpleados();
 parent.opener.opener.location.reload();
 parent.opener.close(); 
 //parent=03 opener=02 parent=00
</script>
<%else%>
<script> 
 parent.cambiar('<%= l_porc%>','<%= trim(l_total_emp) %>','<%= trim(l_restantes_emp) %>','<%= convAHora(l_tiempo_actual)%>','<%= convAHora(l_tiempo_restante) %>','<%= trim(l_sql) %>');
 setTimeout("refrescar()", 3000);
</script>
<%end if%>

</body>
</html>





