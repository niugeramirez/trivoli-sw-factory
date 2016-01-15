<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<% 
Dim l_proximo
Dim l_anterior
Dim l_usuario
Dim l_descripcion
Dim l_tipo
Dim l_codigo
Dim l_esFin
Dim l_esModif
Dim l_esPrimero
Dim l_PuedeVer
Dim l_Secuencia
Dim l_obj
Dim l_sql
Dim l_rs

l_tipo        = request("tipo")
l_descripcion = request("descripcion")
l_codigo      = request("codigo")
l_usuario     = session("UserName") 
l_Secuencia   = 0

if l_tipo = "" then
  l_tipo = "0"
end if

if l_codigo = "" then
  l_codigo = "@@@"
end if

%>

<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Circuíto de firmas - Ticket</title>
</head>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_ay_generica.js"></script>
<script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet") 'lo creo aqui para poder usarlo en las consultas de abajo

l_sql = "select count(*) from cysfincirc "
l_sql = l_sql & "where cysfincirc.cystipnro = " & l_tipo & " and cysfincirc.userid = '" & l_usuario & "'" 

rsOpen l_rs, cn, l_sql, 0 

if l_rs(0) = "0" then
  l_esfin = "false"
else  
  l_esfin = "true"
end if  

l_rs.close

l_sql = "select * from cystipo "
l_sql = l_sql & "where (cystipo.cystipact = -1) and cystipo.cystipnro = " & l_tipo 

rsOpen l_rs, cn, l_sql, 0 

l_PuedeVer = not l_rs.eof

l_rs.close

if l_PuedeVer = false then  'verifico que esten activas las firmas
%>
  alert("Este tipo de firmas electrónicas esta desactivado.");
  window.close()
<% 
end if

l_sql = "select cysfirautoriza, cysfirsecuencia, cysfirdestino from cysfirmas "
l_sql = l_sql & "where cysfirmas.cystipnro = " & l_tipo & " and cysfirmas.cysfircodext = '" & l_codigo & "' " 
l_sql = l_sql & "order by cysfirsecuencia desc"

rsOpen l_rs, cn, l_sql, 0 

l_PuedeVer = true

if not l_rs.eof then
  l_esPrimero = "false"
  l_secuencia = l_rs("cysfirsecuencia")
  if l_rs("cysfirautoriza") = l_usuario then   'Es una modificación del ultimo
    l_proximo = l_rs("cysfirdestino")          'Guardo los datos del proximo 
	l_esModif = "true"
	l_PuedeVer = "true"
    if not l_rs.eof then
      l_rs.movenext
 	  l_anterior = l_rs("cysfirautoriza")
    end if
  else
    if l_rs("cysfirdestino") = l_usuario then
	  l_PuedeVer = "true"
	else
	  l_PuedeVer = "false"
	end if
	l_esModif = "false"
	l_anterior = l_rs("cysfirautoriza")
  end if
else
  l_esPrimero = "true"
  l_esModif   = "false"
end if  

l_rs.close

if l_PuedeVer = "false" then
%>
  alert("No esta autorizado para ver este registro");
  window.close()
<% 
end if
%>

function Validar_Formulario()
{
  if ((document.datos.codproximo.value == "") && 
      ( ! <%= l_esfin %>))
	  alert("Debe ingresar un Empleado.");
  else
    if ((document.datos.codproximo.value == document.datos.codanterior.value) &&
        (document.datos.codproximo.value != "")) 
  	    alert("La próxima autorización es igual a la persona que autorizó anteriormente.");
    else
      if (document.datos.codproximo.value == document.datos.codactual.value)
  	      alert("La próxima autorización es igual a la autorización actual.");
      else
        {
	      document.datos.submit();
        }
}

window.resizeTo(422,205);
</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="Admin_Firmas_03.asp" method="post">
<input type="Hidden" name="tipo" value="<%= l_tipo %>">
<input type="Hidden" name="codigo" value="<%= l_codigo %>">
<input type="Hidden" name="esfin" value="<%= l_esfin %>">
<input type="Hidden" name="esmodif" value="<%= l_esmodif %>">
<input type="Hidden" name="esprimero" value="<%= l_esprimero %>">
<input type="Hidden" name="secuencia" value="<%= l_secuencia %>">
<input type="Hidden" name="descripcion" value="<%= l_descripcion %>">

<table cellspacing="1" cellpadding="0" border="0" width="100%">
  <tr>
    <td class="th2" colspan="6">Circuito de firmas</td>
  </tr>
<tr>
    <td><b>Anteriormente autorizado por:</b><br>
	<input type="text" name="codanterior" size="10" readonly="true" maxlength="20" onchange="javascript:verificacodigochar(this,document.datos.desanterior,'iduser','usrnombre','user_per')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="Text" name="desanterior" size="40" disabled>
	<script>
		document.datos.codanterior.value = "<%= l_anterior %>";
		document.datos.codanterior.onchange();
	</script>
	</td>
</tr>

<tr>
    <td><b>Autorizado por:</b><br>
	<input type="text" name="codactual" size="10" readonly="true" maxlength="20" onchange="javascript:verificacodigochar(this,document.datos.desactual,'iduser','usrnombre','user_per')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="Text" name="desactual" size="40" disabled>
	<script>
		document.datos.codactual.value = "<%= l_usuario %>";
		document.datos.codactual.onchange();
	</script>
	</td>
</tr>

<tr>
    <td><b>Proxima autorización:</b><br>
	<input type="text" name="codproximo" size="10" <% If l_esfin = "true" then %>readonly="true" <% End If %>maxlength="20" onchange="javascript:verificacodigochar(this,document.datos.desproxima,'iduser','usrnombre','user_per')">
	<input type="Button" <% If l_esfin = "true" then %>disabled <% End If %>name="boton" value="?" onclick="javascript:ayudacodigochar(document.datos.codproximo,document.datos.desproxima,'iduser','usrnombre','user_per','','Usuario;Nombre','Usuarios')">
	<input type="Text" name="desproxima" size="40" disabled>
	<script>
		document.datos.codproximo.value = "<%= l_proximo %>";
		document.datos.codproximo.onchange();
	</script>

	</td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</form>
</body>
</html>
