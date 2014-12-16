<html>
<head>
<link href="<%= session("estilo")%>" rel="StyleSheet" type="text/css">
<title>Untitled Document</title>
</head>
<body>
<form name="abmpersona" action="personal_01.asp" target="_self" method="post">
	<% 	If request.Form("abm") = 0 Then 
			consulta()
		else
			abm()
		end if	%>
</form>

<% sub consulta() %>
<input type="hidden" name="abm" value="1">
	<table cellpadding="0" cellspacing="0">
		<tr>
			<th colspan="2" nowrap>Personales <a href="javascript:document.abmpersona.submit();">Modificar</a></th>
		</tr>
			<td>
				<img src="fotos/Hombre7.bmp" class="imgtercero">
			</td>
			<td  width="100%">
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td align="right"nowrap>Nombre :</td>
						<td width="100%"><input class="deshabinp" readonly name="" type="text" value="Juan José"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Apellido :</td>
						<td><input class="deshabinp"  name=""  readonly type="text" value="Perez"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Fecha de Nacimiento :</td>
						<td><input class="deshabinp"  name=""  readonly type="text" value="12/05/1970"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Nacionalidad :</td>
						<td><input class="deshabinp"  name="" readonly  type="text" value="Argentino"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Estado Civil :</td>
						<td><input class="deshabinp"  name="" readonly  type="text" value="Soltero"></td>
					</tr>
				</table>
			</td>
		</table>
<% end sub %>

<% sub abm() %>
	<input type="hidden" name="abm" value="0">
	<table cellpadding="0" cellspacing="0">
		<tr>
			<th colspan="2" nowrap>Personales <a href="javascript:document.abmpersona.submit();">Aceptar</a></th>
		</tr>
			<td>
				<img src="fotos/Hombre7.bmp" class="imgtercero">
			</td>
			<td width="100%">
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td align="right"nowrap>Nombre :</td>
						<td width="100%"><input class="habinp" name="aa" type="text" value="Juan José"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Apellido :</td>
						<td><input class="habinp" name="" type="text" value="Perez"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Fecha de Nacimiento :</td>
						<td><input class="habinp" name="" type="text" value="12/05/1970"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Nacionalidad :</td>
						<td><input class="habinp" name="" type="text" value="Argentino"></td>
					</tr>
					<tr>
						<td align="right" nowrap>Estado Civil :</td>
						<td><input class="habinp" name="" type="text" value="Soltero"></td>
					</tr>
				</table>
			</td>
		</table>
		<script>
			document.abmpersona.aa.focus();
		</script>
<% end sub %>
</body>
<script>
		parent.document.all.abmpersonal.style.height = document.body.scrollHeight;
</script>
</html>
