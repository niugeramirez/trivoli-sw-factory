<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<%
'Modificado     :15-03-2006 Maximiliano Breglia Solo vean empleados activos
on error goto 0

const c_tipoEstr = 1

Dim l_rs
Dim l_sql
Dim l_emplegActual
Dim l_empleado
Dim l_ternro
Dim l_email
Dim l_estructura
Dim l_tipoestr
Dim l_doc
Dim l_sigla

Dim l_primero
Dim l_ultimo
Dim l_anterior
Dim l_siguiente

Dim	l_tipimdire		
Dim	l_tipimanchodef	
Dim	l_tipimaltodef	
Dim	l_terimnombre	

l_emplegActual = request("empleg")

sub EmpleadoSinSeleccionar()

       'Me posiciono en el primer registro
       l_rs.moveFirst

	   'Busco el anterior
	   l_anterior = l_rs("empleg")

	   'El empleado actual
       l_emplegActual = l_rs("empleg")
	   l_empleado = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " &l_rs("ternom2") 
	   l_ternro = l_rs("ternro")
	   l_email  = l_rs("empemail")

	   'Busco el siguiente
	   l_rs.moveNext
	   if l_rs.eof then
 	      l_siguiente = l_anterior
	   else
 	      l_siguiente = l_rs("empleg")
	   end if

end sub 'EmpleadoSinSeleccionar()

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT empemail,ternro,empleg,terape,terape2,ternom,ternom2 FROM empleado WHERE empleado.empest = -1 and empreporta IN (SELECT ternro FROM empleado WHERE empleg=" & Session("empleg") & ")"
l_sql = l_sql & " ORDER BY empleg ASC"

rsOpenCursor l_rs, cn, l_sql, 0 , adOpenDynamic

l_primero   = 0
l_ultimo    = 0
l_anterior  = 0
l_siguiente = 0
l_ternro    = ""
l_email     = ""

if not l_rs.eof then
   'Busco el primero
   l_rs.moveFirst
   l_primero = l_rs("empleg")

   'Busco el ultimo
   l_rs.moveLast
   l_ultimo = l_rs("empleg")
   
   if l_emplegActual = "" then
   
       call EmpleadoSinSeleccionar()

   else

       'Me posiciono en el primer registro
       l_rs.moveFirst

       'Busco el legajo
       l_rs.find("empleg=" & l_emplegActual)
	   
	   if l_rs.eof then
	   
          call EmpleadoSinSeleccionar()

	   else

		   'El empleado actual
	       l_emplegActual = l_rs("empleg")
		   l_empleado = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " &l_rs("ternom2") 
		   l_ternro = l_rs("ternro")
		   l_email  = l_rs("empemail")
		   
		   'Busco el anterior
		   l_rs.movePrevious()
		   if l_rs.bof then
		      l_anterior = l_emplegActual
		   else
		      l_anterior = l_rs("empleg")
		   end if
	
		   'Busco el siguiente
		   l_rs.moveNext
		   l_rs.moveNext
		   if l_rs.eof then
	 	      l_siguiente = l_emplegActual
		   else
	 	      l_siguiente = l_rs("empleg")
		   end if

		end if
   
   end if
   
end if

l_rs.close

if l_ternro <> "" then
   'Busco el documento del empleado
   l_sql = " SELECT nrodoc,tidsigla FROM ter_doc docu "
   l_sql = l_sql & "INNER JOIN tipodocu  ON tipodocu.tidnro= docu.tidnro AND docu.tidnro>0 and docu.tidnro<5 "
   l_sql = l_sql & "WHERE docu.ternro= " & l_ternro
   
   rsOpen l_rs, cn, l_sql, 0 
   
   if not l_rs.eof then
      l_doc   = l_rs("nrodoc")
	  l_sigla = l_rs("tidsigla")
   end if
   
   l_rs.close

   'Busco una estructura de la empresa
   l_sql = " SELECT tipoestructura.tedabr "
   l_sql = l_sql & " FROM tipoestructura WHERE tipoestructura.tenro = " & c_tipoEstr

   rsOpen l_rs, cn, l_sql, 0 
   
   if not l_rs.eof then
	  l_tipoestr   = l_rs("tedabr")
   end if
   
   l_rs.close
   
   'Busco una estructura de la empresa
   l_sql = " SELECT estructura.estrdabr "
   l_sql = l_sql & " FROM his_estructura "
   l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND his_estructura.tenro=" & c_tipoEstr & " AND his_estructura.ternro=" & l_ternro & " AND his_estructura.htetdesde <= " & cambiafecha(date(),"YMD",true) & " AND ((" & cambiafecha(date(),"YMD",true) & " <= his_estructura.htethasta) OR (his_estructura.htethasta IS NULL)) "

   rsOpen l_rs, cn, l_sql, 0 
   
   if not l_rs.eof then
      l_estructura = l_rs("estrdabr")
   end if
   
   l_rs.close
   
end if

%>
<html>
<head>
<link href="<%= c_estilo %>" rel="StyleSheet" type="text/css">
<title>Untitled Document</title>
</head>
<script>
function cambiarA(legajo){
  window.location = "terceros.asp?empleg=" + legajo;
}
</script>
<body>
	<table cellpadding="0" cellspacing="0" class="Tterceros" width="100%">
		<tr>
			<td class="tdimgtercero" style="width:100px;">
			<%
			l_sql = "SELECT * "
			l_sql = l_sql & " FROM tipoimag "
			l_sql = l_sql & " WHERE tipimnro=3 "

			rsOpen l_rs, cn, l_sql, 0
			
			if l_rs.eof then
				l_tipimanchodef	= 50
				l_tipimaltodef	= 50
				l_tipimdire		= "/serviciolocal/fotos"
			else
				l_tipimanchodef	= l_rs("tipimanchodef")
				l_tipimaltodef	= l_rs("tipimaltodef")
				l_tipimdire		= l_rs("tipimdire")				
			end if
			
			l_rs.close

			l_sql = "SELECT terimnombre, ter_imag.terimfecha "
			l_sql = l_sql & " FROM ter_imag "
			l_sql = l_sql & " WHERE ter_imag.ternro = " & l_ternro
			l_sql = l_sql & " ORDER BY ter_imag.terimfecha DESC "
			
			rsOpen l_rs, cn, l_sql, 0
			
			if not l_rs.eof then
				l_terimnombre	= l_rs("terimnombre")
	
				if trim(l_terimnombre) <> "" then %>		
				<img src="<%=l_tipimdire%><%=l_terimnombre%>" height="<%=l_tipimaltodef%>" width="<%=l_tipimanchodef%>" alt="" border="1" class="imgtercero">		
				<% else %>		
				<img height="<%=l_tipimaltodef%>" width="<%=l_tipimanchodef%>" src="/serviciolocal/shared/fotos/nofoto.jpg" alt="" border="0" class="imgtercero">		
				<%end if  
			else%>
				<img height="<%=l_tipimaltodef%>" width="<%=l_tipimanchodef%>"  src="/serviciolocal/shared/fotos/nofoto.jpg" alt="" border="0" class="imgtercero">		
			<%end if 
			l_rs.Close
	         %>		
			</td>
			<td class="tdtercero" align="left"> <!--  width="75%"-->
				<table cellpadding="0" cellspacing="0" border="0" width="100%" align="left">
					<tr>
						<td class="tdterceroTitulo" width="1%">Legajo : </td>
						<td class="tdtercerotext">
							<a href="Javascript:cambiarA(<%= l_primero%>);"   alt="<%= l_primero%>" class="flecha"><img src="shared/images/ffirst.gif" class="imgflechas"></a>
							<a href="Javascript:cambiarA(<%= l_anterior%>);"  alt="<%= l_anterior%>" class="flecha"><img src="shared/images/fprev.gif" class="imgflechas"></a>
							<input type="text" value="<%= l_emplegActual%>" onchange="cambiarA(this.value)" size="10">
							<a href="Javascript:cambiarA(<%= l_siguiente%>);" alt="<%= l_siguiente%>" class="flecha"><img src="shared/images/fnext.gif" class="imgflechas"></a>
							<a href="Javascript:cambiarA(<%= l_ultimo%>);"    alt="<%= l_ultimo%>" class="flecha"><img src="shared/images/flast.gif" class="imgflechas"></a>
							
						</td>
					</tr>
					<tr>
						<td class="tdterceroTitulo" width="1%">Nombre : </td>
						<td class="tdtercerotext"><%= l_empleado%></td>
					</tr>
					<tr>
						<td class="tdterceroTitulo" width="1%"><%= l_sigla%> : </td>
						<td class="tdtercerotext"><%= l_doc%></td>
					</tr>
					<tr>
						<td class="tdterceroTitulo" width="1%">E-Mail : </td>
						<td class="tdtercerotext"><a href="mailto:<%= l_email%>" class="mail"><%= l_email%></a></td>
					</tr>
					<tr>
						<td class="tdterceroTitulo" width="1%"><%= l_tipoestr%> : </td>
						<td class="tdtercerotext"><%= l_estructura%></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</body>
</html>
<%
set l_rs = Nothing

cn.Close
set cn = Nothing
%>
<script>
  parent.accionTercero(<%= l_emplegActual%>);
</script>
