<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<html>
<head>
<link href="<%= c_estilo%>" rel="StyleSheet" type="text/css">
<title>Untitled Document</title>
</head>

<script>
function aceptar(){
   var errores = 0;
   
   if (document.datos.docu.value == ""){
       alert('Debe ingresar un número de documento.')
	   errores++;
   }

   if ((errores==0) && (document.datos.pass1.value != document.datos.pass2.value)){
       alert('Las contraseñas deben ser iguales.')
	   errores++;
   }
   
   if (errores == 0){
       document.datos.submit();
   }

}
</script>
<body>

<form name="datos" action="RegistroBD.asp" method="post">
<table width="100%">
	<tr>
		<td class="barmenu">
			Recursos Humanos
		</td>
	</tr>
	<tr>
		<th colspan="2">Registro de usuarios.</th>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table>  
				<tr>
					<td align="right">Nro. documento :</td>
					<td><input name="docu" type="text" size="15"></td>
				</tr>
				<tr>
					<td align="right">Contraseña :</td>
					<td><input name="pass1" type="password" size="15"></td>
				</tr>
				<tr>
					<td align="right">Confirmar contraseña :</td>
					<td><input name="pass2" type="password" size="15"></td>
				</tr>
				<tr>
					<td colspan="2" align="right">
						<a href="javascript:aceptar();" class="sidebtnSHW">Aceptar</a>
					</td>
				</tr>
			</table>
		</td>
		<tr>
			<td>
			<br>
			<ul>
			<li>
			Si Ud. ya se ha registrado, en USUARIO coloque su número de documento y en CONTRASEÑA indique la que Ud. tenga registrada.
			</li>
			<li>
			Si Ud. nunca ingresó a la Auto Gestión, utilice la opción REGISTRARSE para poder indicar su contraseña y quedar habilitado.
			</li>
			<li>
			Si no se puede registrar, deberá comunicarse con su Administrador de RH Pro para que lo habilite para el acceso a Auto Gestión.
			</td>
			</li>
			</ul>
		</tr>
	</tr>
</table>
</form>

</body>
<script>
		parent.document.all.centro.style.height = document.body.scrollHeight;
</script>
</html>
