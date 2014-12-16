<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: balanzas_con_02.asp
'Descripción: ABM de Balanzas
'Autor : Gustavo Manfrin
'Fecha: 19/04/2005
'Modificado: 

on error goto 0


'Datos del formulario
Dim l_balnro
Dim l_baldes
Dim l_balcod
Dim l_balact
Dim l_planro
Dim l_balvpc
Dim	l_balmarca
Dim l_balconexion

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Balanzas - Ticket</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (Trim(document.datos.balcod.value) == ""){
	alert("Debe ingresar el Código.");
	document.datos.balcod.focus();
	}
else if(!stringValido(document.datos.balcod.value)){
	alert("El Código contiene caracteres inválidos.");
	document.datos.balcod.focus();
	}
else if(Trim(document.datos.baldes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.baldes.focus();
	}
else if(!stringValido(document.datos.baldes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.baldes.focus();
	}
else if(Trim(document.datos.planro.value) == ""){
	alert("Debe ingresar una Planta.");
	document.datos.planro.focus();
	}
else{
	var d=document.datos;
	document.valida.location = "balanzas_con_06.asp?tipo=<%= l_tipo%>&balnro="+document.datos.balnro.value + "&balcod="+document.datos.balcod.value + "&baldes="+document.datos.baldes.value ;
	}	
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.balcod.focus();
}

</script>
<% 
select Case l_tipo
	Case "A":
		l_balnro = ""
		l_baldes = ""
		l_balcod = ""
		l_balact = ""
		l_planro = ""
		l_balvpc = ""
		l_balmarca = ""		
		l_balconexion = ""
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_balnro = request.querystring("cabnro")
		
		l_sql = "SELECT  tkt_planta.plades, tkt_balanza.* "
		l_sql = l_sql & " FROM tkt_balanza "
		l_sql = l_sql & " LEFT JOIN tkt_planta ON tkt_balanza.planro= tkt_planta.planro "
		l_sql  = l_sql  & " WHERE balnro = " & l_balnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_baldes = l_rs("baldes")
			l_balcod = l_rs("balcod")
			l_balact = l_rs("balact")
			l_planro = l_rs("planro")
			l_balvpc = l_rs("balvpc")
			l_balmarca = trim(l_rs("balmarca"))
			l_balconexion = trim(l_rs("balconexion"))			
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.planro.focus()">
<form name="datos" action="balanzas_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="balnro" value="<%= l_balnro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Balanzas</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
					    <td align="right"><b>Planta:</b></td>
							<td>
				  		   <select name="planro" style="width:150px;">
							<% If l_tipo = "A" then
							Set l_rs = Server.CreateObject("ADODB.RecordSet")%>
							<option value="">&laquo; Seleccione una opción &raquo;</option>
							<% End If 
							
							l_sql = "SELECT planro,plades"
							l_sql = l_sql & " FROM tkt_planta"
							l_sql = l_sql & " ORDER BY planro"
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("planro")%>"><%=l_rs("plades")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							%>
							<script>
								document.datos.planro.value = "<%=l_planro%>";
							</script>
							</select>
                        </td>
				
					</tr>
					<tr>
					    <td align="right"><b>Código:</b></td>
						<td>
							<input type="text" name="balcod" size="18" maxlength="12" value="<%= l_balcod %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Descripción:</b></td>
						<td>
							<input type="text" name="baldes" size="60" maxlength="50" value="<%= l_baldes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Marca:</b></td>
						<td>
				  		   <select name="balmarca" style="width:150px;">
							<option value="">&laquo; Seleccione una opción &raquo;</option>
  						    <option value="ToledoA">ToledoA</option>
							</select>
            				<script>
								document.datos.balmarca.value = "<%=l_balmarca%>";
							</script>

						</td>
					</tr>
					<tr>
					    <td align="right" nowrap><b>Config. conexión:</b></td>
						<td>
							<input type="text" name="balcon" size="60" maxlength="50" value="<%= l_balconexion %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Activa:</b></td>
					    <td align="left"> <input type="Checkbox" name="balact" <% If l_balact = -1 then  %>checked<% end if %>></td>
					</tr>
					<tr>
					    <td align="right" nowrap><b>Verifica Posición:</b></td>
          			    <td align="left"> <input type="Checkbox" name="balvpc" <% If l_balvpc = -1 then  %>checked<% end if %>></td>
					<tr>
         		</table>
				</td>
				<td width="50%"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
<iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
