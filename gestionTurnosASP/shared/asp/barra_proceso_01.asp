<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
'-----------------------------------------------------------------------------------
Archivo	    : barra_proceso_01.asp
Descripción : Controla el estado del proceso
Autor		: Scarpa D.
Fecha		: 05/02/2004
Modificado	: 
	18-02-04 Favre F. Se standariso
------------------------------------------------------------------------------------
-->
<% 
'Variables base de datos
 Dim l_rs
 Dim l_cm
 Dim l_sql
 
'Variables uso local
 Dim l_porc
 
 Dim l_tiempo_actual
 Dim l_tiempo_restante
 Dim l_tiempo_total
 
 Dim l_bpronro
 Dim l_funcion
 Dim l_parametros
 
 l_bpronro 	  = request("bpronro")
 l_funcion	  = Request.QueryString("funcion")
 l_parametros = Request.QueryString("parametros")
 
%>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<html>
<head>
<title><%= Session("Titulo")%>Estado Proceso</title>
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

rsOpen l_rs, cn, l_sql, 0
if not l_rs.EOF then
   if CStr(l_rs("bprcestado") = "Procesando") or CStr(l_rs("bprcestado") = "Pendiente") then

	   if isNull(l_rs("bprcprogreso")) then
	      l_porc = "0"   
		  l_tiempo_actual = "0"
	   else
	      if l_rs("bprcprogreso") = "" then
		     l_porc = "0"
			 l_tiempo_actual = "0"
		  else
		     l_porc = CInt(l_rs("bprcprogreso"))
			 if l_rs("bprcprogreso") <> "" then
			 	if isnull(l_rs("bprctiempo")) then
					l_tiempo_actual = 0
				else
				    l_tiempo_actual = CLng(l_rs("bprctiempo"))
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
   l_tiempo_total = (CLng(l_tiempo_actual) * 100) / CInt(l_porc)
   l_tiempo_restante = l_tiempo_total - l_tiempo_actual
else
   l_tiempo_total = 0
   l_tiempo_restante = 0
end if

%>

<script>
function refrescar(){
  window.location = "barra_proceso_01.asp?bpronro=<%= l_bpronro%>&parametros=<%= l_parametros%>&funcion=<%= l_funcion %>";
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

<%if CInt(l_porc) = 100 then%>
<script>
 parent.cambiar('<%= l_porc%>','<%= convAHora(l_tiempo_actual)%>','<%= convAHora(l_tiempo_restante) %>','<%= l_sql%>');
 parent.parent.<%= l_funcion %>(<%= l_parametros%>);
</script>
<%else%>
<script> 
 parent.cambiar('<%= l_porc%>','<%= convAHora(l_tiempo_actual)%>','<%= convAHora(l_tiempo_restante) %>','<%= l_sql%>');
 setTimeout("refrescar()", 3000);
</script>
<%end if%>

</body>
</html>





