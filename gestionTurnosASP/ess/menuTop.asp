<% Option Explicit 
'response.write "flor"
'response.end	
%>

<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<%

'modificado: 13/08/2007 - Lisandro moro - se quito la variable ternro pasada por querystring.
on error goto 0



Dim l_rs
Dim l_sql

Dim l_accion
Dim l_orden
Dim l_empleado
Dim l_raiz
Dim l_perfil
Dim l_empleg
Dim l_menu

l_accion = request("src")
l_orden  = request("orden")
l_empleg = request("empleg")
l_menu   = request("menu")




Set l_rs = Server.CreateObject("ADODB.RecordSet")	

if Trim(l_accion) = "" then
   l_accion = "bienvenida.asp"
end if

if Trim(l_menu) = "" then
   l_menu = "ess"
end if

function Habilitado(acceso,perfil)

   if acceso = "*" then
       Habilitado = true
   else
       Habilitado = (instr(Trim(acceso),Trim(perfil)) > 0)
   end if

end function 'Habilitado(acceso,perfil)

'Busco los datos del empleado

'l_sql = "SELECT * FROM empleado WHERE empleg=" & Session("empleg")

'rsOpen l_rs, cn, l_sql, 0



l_empleado = ""

'if not l_rs.eof then
'   l_empleado = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " &l_rs("ternom2") 
'end if

'l_rs.close

%>
<html>
<head>
<link href="<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Principal</title>
</head>
<script src="shared/js/fn_util.js"></script>
<script>
var jsSelRow = null;
var jsSelMenuOrden = null;
var accionMenu='#';

function Deseleccionar(boton){
	boton.className = "boton";
}

function Seleccionar(boton,orden){
	if (jsSelRow != null){
	   Deseleccionar(jsSelRow);
	}
	boton.className = "botonsel";
	jsSelRow = boton;
    jsSelMenuOrden = orden;	
}

function accion(menu,orden,accio){
    Seleccionar(menu,orden);
    parent.accion(<%= l_orden%>,accio);
	accionMenu = accio;
}

function accion2(menu,orden){
    Seleccionar(menu,orden);
    parent.accion(<%= l_orden%>,accionMenu);
}

function ira(accio)
{
   accionMenu = accio;
}

</script>
<body class="indexprincipal" >
	<table class="Tprincipal" cellpadding="0" cellspacing="0">
 		<tr>
			<td class="barmenu" ><%= l_empleado%>
			<%
			'Busco cual es el nro de raiz del menu
			
			l_sql = "SELECT menunro FROM menuraiz where upper(menudesc) = upper('" & l_menu & "')"

			rsOpen l_rs, cn, l_sql, 0
	
			l_raiz = 0
			if not l_rs.eof then
			   l_raiz = l_rs("menunro")
			end if
	
			l_rs.close
			
			'Busco el perfil del empleado
'			l_sql = "SELECT perfnom FROM empleado "
'			l_sql = l_sql & "INNER JOIN perf_usr ON empleado.perfnro = perf_usr.perfnro WHERE empleg=" & Session("empleg")	 
			 
'		    rsOpen l_rs, cn, l_sql, 0 
			 
'			if l_rs.eof then
			    l_perfil = "#NO TIENE#"
'			else
'			    l_perfil = l_rs("perfnom")
'			end if

'			l_rs.close
			
			'Busco los elementos del menu
			if Trim(l_orden) <> "" then
				l_sql = "SELECT menuaccess,menuname, parent, menuorder, action FROM menumstr where menuraiz = " & l_raiz
				l_sql = l_sql & " AND upper(parent) LIKE upper('" & l_orden & l_menu & "') "
				l_sql = l_sql & " ORDER BY menuorder"
				
				rsOpen l_rs, cn, l_sql, 0 

				do until l_rs.eof
				   'if Habilitado(l_rs("menuaccess"),l_perfil) then
				      if inStr(UCASE(l_rs("action")),"JAVASCRIPT") > 0 then
						%><a href="#" class="boton" onClick="<%= l_rs("action")%>;alert();accion2(this,'<%= l_rs("menuorder")%>');">7<%= l_rs("menuname")%></a><% 
					  else
						%><a href="#" class="boton" onClick="Javascript:accion(this,'<%= l_rs("menuorder")%>','<%= l_rs("action")%>');"><%= l_rs("menuname")%></a><% 
					  end if
				   'end if
				
				   l_rs.movenext
				loop
					%>
						<span class="spanmenutop">&nbsp;</span>
					<%
				l_rs.close
			end if
			%>	
			</td>
		</tr>
		<tr>
			<td class="tdprincipal">
				<% 
				If l_accion = "" or l_accion = "#" Then 
					'l_accion = "blanc.html"
					l_accion = "blanc.asp"
				end if
								
				%>
				<iframe scrolling="No" name="principal" src="<%= l_accion%>" frameborder="0" ></iframe>
			</td>
		</tr>
	</table>
</body>
<script>
//		parent.document.all.centro.height = document.body.scrollHeight;
</script>
</html>

<%

  set l_rs = nothing
  
  cn.close
  set cn = nothing
%>

