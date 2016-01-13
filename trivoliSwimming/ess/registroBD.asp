<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/encrypt.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/emails.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<html>
<head>
<link href="<%= c_estilo%>" rel="StyleSheet" type="text/css">
<title>Autogestión</title>
</head>
<script>
function volver(){
    parent.document.location.reload();
//  parent.document.all.ifrmmenu.location = "menu.asp";
//  document.location = "principal.asp";
}
</script>
<body class="indexmenu">
<%
on error goto 0

Dim l_sql
Dim l_rs
Dim l_cm

Dim l_docu
Dim l_pass
Dim l_passNuevo
Dim l_passNuevoCrypt

Dim	l_asunto 
Dim	l_msgBody
Dim	l_destino

Set l_rs = Server.CreateObject("ADODB.RecordSet")	

l_pass = request.form("pass1")
l_docu = request.form("docu")

sub mensajes(nro)
%>
 		 <table cellpadding="0" cellspacing="0"  border="0" class="Tmenu">
			<tr>
				<td class="barmenu">
					Recursos Humanos
				</td>
			</tr>
				<tr>
					<th colspan="2">Registro de usuarios.</th>
				</tr>
			<tr>
				<td class="__menumenu">
					<div class="menumenu">
					<%select case nro%> 
					   <%case 1%>
					   <div class="menutextoerror">					   
					   El nro. de documento no es valido.
				  	   </div>					   					   
					   <%case 2%>
					   <div class="menutextoerror">					   					   
					   El usuario no tiene un e-mail asociado.
				  	   </div>					   					   
					   <%case 3%>
					   <div class="menutexto">					   					   
					   Se ha enviado a su direcci&oacute;n de e-mail su contraseña.					   
				  	   </div>			
					   <%case 4%>
					   <div class="menutextoerror">					   					   
					   El usuario ingresado ya se encuentra registrado, se ha enviado a su direcci&oacute;n de e-mail su contraseña.
				  	   </div>			
					<%end select%>
					<div class="menutexto"><a class="btnmenu" href="javascript:volver();">Volver</a></div>
					</div>
				</td>
			</tr>
		</table>   
<%
end sub 'mensajes(nro)

Dim l_existe
l_existe = false

l_sql = "SELECT ter_doc.ternro, empleado.emppass,empleado.empemail, empleado.empleg, empleado.empessactivo FROM ter_doc, empleado "
l_sql = l_sql & " WHERE empleado.ternro = ter_doc.ternro "
l_sql = l_sql & " AND ter_doc.nrodoc = '" & l_docu & "'"
l_sql = l_sql & " AND ter_doc.tidnro " 
if ( c_TipoDocAg <= 4 ) then
	l_sql = l_sql & " <= 4 " 
else
	l_sql = l_sql & " = " & c_TipoDocAg
end if

rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
   if isNull(l_rs("empemail")) then
	   l_rs.close
	   call mensajes(2)
   else
	   if not isNull(l_rs("emppass")) then
	      'Le envio por mail el password actual
	      l_passNuevo = Decrypt(c_seed1,l_rs("emppass"))
		  l_existe = true
	   else
	      if Trim(l_pass) = "" then
		     'si es vacio genero un password nuevo
			 l_passNuevo = mid(genClaveUnica,1,10)
		  else
		     'le envio el password que ingreso
			 l_passNuevo = l_pass
		  end if
	   end if
	
	    l_passNuevoCrypt = Encrypt(c_seed1,l_passNuevo)
	
	    set l_cm = Server.CreateObject("ADODB.Command")
		
		l_sql = "UPDATE empleado SET emppass = '" & l_passNuevoCrypt
		l_sql = l_sql & "' WHERE empleado.ternro = " & l_rs("ternro")
	
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		Set l_cm = Nothing
		
		l_asunto      = "Contraseña Autogestion RH Pro X2"
		l_msgBody     = "Su contraseña es: " & l_passNuevo
		l_destino     = l_rs("empemail")
		
		l_rs.close
		
		generarMail "",l_asunto,l_msgBody,l_destino	
		
		if l_existe then
			call mensajes(4)
		else
			call mensajes(3)
		end if
   end if
else
   l_rs.close
   call mensajes(1)
end if

set l_rs = nothing
cn.close
set cn = nothing
%>
</body>
</html>