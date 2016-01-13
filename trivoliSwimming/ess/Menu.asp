<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->

<%

on error goto 0

Dim l_usuario
Dim l_password

Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_perfil
Dim l_primero
Dim l_tipo_menu
Dim l_baseActual

l_usuario   = request.form("usuario")
l_password  = request.form("password")
l_tipo_menu = request("menu")
l_baseActual = request("base")

if len(trim(l_baseActual)) = 0 then
	l_baseActual = 2 'Session("base")
end if

Session("base") = 2 'l_baseActual

'response.write "base " & Session("base") & "<br>"
%>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

'Controlo que tipo de menu tiene que usar
'if Trim(l_tipo_menu) = "" then
   l_tipo_menu = "ess"
'else
'   if UCase(l_tipo_menu) = "MSS" then
'      if not tieneReportaA() then
' 	     l_tipo_menu = "ESS"
'	  end if
'   end if
'end if

function Habilitado(acceso,perfil)

   if acceso = "*" then
       Habilitado = true
   else
       Habilitado = (instr(Trim(acceso),Trim(perfil)) > 0)
   end if


end function 'Habilitado(acceso,perfil)

'Function Decrypt(ByVal sClave, ByVal sOriginal, blnAccion)
Function Decrypt(sClave, sOriginal, blnAccion)
Dim LenOri 
Dim LenClave
Dim i, j 
Dim cO, cC 
Dim k 
Dim v 

LenOri = Len(sOriginal)
LenClave = Len(sClave)

v = ""
i = 0
For j = 1 To LenOri
    i = i + 1
    If i > LenClave Then
        i = 1
    End If
    cO = Asc(Mid(sOriginal, j, 1))
    cC = Asc(Mid(sClave, i, 1))
    If blnAccion Then
        k = cO + cC
        If k > 255 Then
            k = k - 255
        End If
    Else
        k = cO - cC
        If k < 0 Then
            k = k + 255
        End If
    End If
    v = v & Chr(k)
Next

'response.write "rrrrrrrrrrrrrr===  " & v
'response.end

Decrypt = v
End Function


%>
<html>
<head>
<link href="<%= c_estilo%>" rel="StyleSheet" type="text/css">
<title>Autogestión</title>
<script src="shared/js/fn_util.js"></script>
<script>
var jsSelRow = null;
var jsSelMenuOrden = null;
var accionMenu='#';

function Deseleccionar(boton){
	boton.className = "btnmenu";
}

function Seleccionar(boton,orden){
	if (jsSelRow != null){
		Deseleccionar(jsSelRow);
	 }
	boton.className = "btnmenusel";
	jsSelRow = boton;
    jsSelMenuOrden = orden;	
}

function accion(menu,orden,accio){
//alert(menu);
//alert(orden);
//alert(accio);
    Seleccionar(menu,orden);
    parent.accion(orden,accio,'<%= l_tipo_menu%>');
	accionMenu = accio;
}

function accion2(menu,orden){
    Seleccionar(menu,orden);
    parent.accion(orden,accionMenu,'<%= l_tipo_menu%>');
}

function salir()
{
   if (confirm('¿ Desea salir de Autogestión Recursos Humanos ?') == true){
    	parent.location = "closesession.asp";
   }
}

function ira(accio)
{
   accionMenu = accio;
}

function cambiarAMenu(tipoMenu){
   parent.cambioMSS();
   accionMenu = '';
   window.location = "menu.asp?menu=" + tipoMenu;
}

</script>
</head>

<body class="indexmenu" onLoad="if (parent.ajustarIframe) parent.ajustarIframe(window);">
<form name="menu" action="Menu.asp" method="post">
	<% 


	
'	    if Trim(session("empleg")) = "" then
		    'Si hay un tipo de documento definido entonces usa la seguridad integrada de windows
	    	'if c_SegIntegradaTipoDoc <> 0 then
			   'if validaUsuarioSeguridadIntegrada() then
			    '  call menu()
			   'else
			    '  call mensajes(1)
			   'end if
			'else
			  'sino pide un usuario y contraseña para entrar
			   if Trim(l_usuario) <> "" then
			   
			   		'response.write "----- " & validaUsuario(l_usuario,l_password)
					'response.end
			   
				   select case validaUsuario(l_usuario,l_password)
				      case 0
			            call loguin(1)
				      case 1
			            call mensajes(2)
					  case 2
			            call loguin(2)
				      case 3
				        call menu()
				   end select
			   else
		    	   call loguin(0)
			   end if
	  	    'end if
'		else
'	        call menu()
'		end if

%>
</form>

<% sub imagenMenu(imagen) %>
	<img class="imgmenu" src="shared/images/a_<%= imagen %>.gif">
<% end sub %>

<% sub loguin(mostrarMensaje) %>
 	<!--<div class="Tmenu"> -->
 		 <table cellpadding="0" cellspacing="0"  border="0" class="Tmenu">
			<tr>
				<td class="barmenu">
					Ingreso
				</td>
			</tr>
			<tr>
				<!-- <td class="__menumenu"> ** -->
				<td class="menumenu">
					<div class="menumenu">
					<div class="menutexto">
					<input type="Hidden" class="inputmenu" name="base" value="2" type="text" size="14" >							
					</div>
					<div class="menutexto">Usuario:</div>
					<div class="menutexto"><input class="inputmenu" name="usuario" type="text" size="14" ></div>
					<div class="menutexto">Contraseña:</div>
					<div class="menutexto"><input class="inputmenu" name="password" type="password" size="14"></div>					
					<%select case mostrarMensaje%> 
					<%case 1%>					
					<div class="menutextoerror">Acceso Invalido</div>					
					<%case 2%>					
					<div class="menutextoerror">Contraseña Invalida</div>										
					<%end select%>					
					<div class="menutexto"><a class="btnmenu" href="javascript:document.menu.submit();">Aceptar</a></div>
					
					<%'if c_PassAutogestion <> 0 then%> 
					<!--
					<div class="menutexto"><a class="btnmenu" href="#" onClick="javascript:parent.document.all.centro.src = 'registro.asp';">Registrarse</a></div>
					-->
					<%'end if%>
					</div>
				</td>
			</tr>						
		</table> 

<% end sub %>

<% sub menu() %>

 		 <table cellpadding="0" cellspacing="0"  border="0" class="Tmenu">
			<tr>
				<td class="barmenu">Menú Principal</td>
			</tr>
			<tr>
				<td>
		<div class="menumenu">
		<%
		Dim l_raiz
		
		'Busco cual es el nro de raiz del menu
		l_sql = "SELECT menunro FROM menuraiz where upper(menudesc) = '" & Ucase(l_tipo_menu) & "'"
		
		rsOpen l_rs, cn, l_sql, 0

		l_raiz = 0
		if not l_rs.eof then
		   l_raiz = l_rs("menunro")
		end if

'	    Response.write "<script>alert('" & l_raiz & "');</script>"
'	    Response.write "<script>alert( '" & Session("empleg") & "');</script>"
		
'	    response.write l_raiz & Ucase(l_tipo_menu)
		
		l_rs.close
		
		'Busco el perfil del empleado
		l_sql = "SELECT perfnom FROM buq_usuario "
		l_sql = l_sql & "INNER JOIN perf_usr ON buq_usuario.perfnro = perf_usr.perfnro WHERE userid= '" & Session("empleg") & "'"	 

	    rsOpen l_rs, cn, l_sql, 0 

 
		if l_rs.eof then
		    l_perfil = "#NO TIENE#"
		else
		    l_perfil = l_rs("perfnom")
		end if
		l_rs.close
		
'	    Response.write "<script>alert('" & l_perfil & "');</script>"
		

		'Busco los elementos del menu
		l_sql = "SELECT menuaccess,menuname, parent, menuorder, action FROM menumstr where menuraiz = " & l_raiz
		l_sql = l_sql & " AND upper(parent) = '" & Ucase(l_tipo_menu) & "' "
		l_sql = l_sql & " ORDER BY menuorder"
		
		'response.write l_sql
		
		rsOpen l_rs, cn, l_sql, 0
		
		l_primero = true

		do until l_rs.eof
		'response.write l_rs("menuname")
' 		   if (l_primero AND tieneReportaA()) OR (not l_primero) then
			   if Habilitado(l_rs("menuaccess"),l_perfil) then

			      if inStr(UCASE(l_rs("action")),"JAVASCRIPT") > 0 then
			  		    'response.write "si"

					%>
					<a href="#" class="btnmenu" onClick="<%= l_rs("action")%>;accion2(this,'<%= l_rs("menuorder")%>');"><li class="limenu"><%= l_rs("menuname")%></li></a>
					<% 
				  else
 			  		    'response.write "no"
					%>
					<a href="#" class="btnmenu" onClick="Javascript:accion(this,'<%= l_rs("menuorder")%>','<%= l_rs("action")%>');"><li class="limenu"><%= l_rs("menuname")%></li></a>
					<% 
				  end if


			   end if
'		   end if

		   l_primero = false

		   l_rs.movenext
		loop
		
		l_rs.close
		%>	

		</div>
				</td>
			</tr>
		</table> 
	<input type="hidden" name="menu" value="0">
<% end sub %>

<% sub mensajes(nro)%>
 		 <table cellpadding="0" cellspacing="0"  border="0" class="Tmenu">
			<tr>
				<td class="barmenu">
					Ingreso
				</td>
			</tr>
			<tr>
				<!-- <td class="__menumenu"> ** -->
				<td class="menumenu">
					<div class="menumenu">
					<div class="menutextoerror">
					<%select case nro%> 
					   <%case 1%>
					   Usuario<br>no<br>autorizado
					   <%case 2%>
					   Usuario sin<br>permiso de acceso
					<%end select%>
					</div>
					</div>
				</td>
			</tr>
		</table>   
<% end sub 'mensajes()%>

<%
  function validaUsuario(l_usuario,l_password)

    Dim l_estado
	Dim l_pass
	Dim l_passDec
	Dim l_legajoUsuario
	
	l_estado = 0
	
	l_sql = "SELECT * FROM buq_usuario "
	l_sql = l_sql & " WHERE buq_usuario.userid = '" & l_usuario & "'"	
	rsOpen l_rs, cn, l_sql, 0 
	
	if not l_rs.eof then
	    ' usuario correcto, me fijo si esta activo o no

'	   Response.write "<script>alert('" & l_rs("userid") & "-" & l_rs("pass") & "');</script>"		
		
       l_estado = 3
       l_legajoUsuario = l_rs("userid")
	   l_pass  = l_rs("pass")
		
	else
		' usuario incorrecto
        l_estado = 0
	end if
	
	l_rs.Close
	
	if l_estado = 3 then
		if isNull(l_pass) then
			' No tiene paswword
	        l_estado = 2
		else
		
			'l_passDec = Decrypt(c_seed1,l_pass, true)
			
			l_passDec = l_pass
			
			'response.write  "---" & c_seed1 & " -- " & l_pass & " -- " & l_passDec
			'response.end						
						
			'l_passDec = Decrypt("1",l_pass)
			
'		    Response.write "<script>alert('" & l_password & "-" & l_passDec & "');</script>"					
			
			
			if (l_password = l_passDec) then
				' Deberia entrar al sistema
				l_estado = 3
				Session("empleg") = l_legajoUsuario
			else
				' password incorrecto
		        l_estado = 2
			end if
		end if
	end if

	validaUsuario = l_estado
  
  end function 'validaUsuario(l_usuario,l_password)
  
  function tieneReportaA()
     
	 l_sql = "SELECT empleg FROM empleado WHERE empreporta IN (SELECT ternro FROM empleado WHERE empleg=" & Session("empleg") & ")"
	 
     rsOpen l_rs2, cn, l_sql, 0 
	 
	 tieneReportaA = not l_rs2.eof
	 
	 l_rs2.close
	 
  end function 'tieneReportaA()  
%>
<script>
//	parent.document.all.ifrmmenu.style.height = document.body.scrollHeight;
</script>
<%

  set l_rs = nothing
  set l_rs2 = nothing  
  
  cn.close
  set cn = nothing
%>
</body>
</html>