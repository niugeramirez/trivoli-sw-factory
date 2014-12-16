<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: mermas_rubros_con_02.asp
'Descripción: ABM de Tipos de Mermas para rubros
'Autor : Alvaro Bayon
'Fecha: 16/02/2005
'Modificado: Raul Chinestra - 16/03/2005 - Solicita obligatoriamente lugar y Rubro para el caso de Alta

'Datos del formulario
Dim l_tipmernro
Dim l_rubdes
Dim l_lugcod
Dim l_lugnro
Dim l_rubnro
Dim l_tipmer
Dim l_forcal


'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Tipos de Mermas para Rubros - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){

<%' if l_tipo="A" then %>
/*	if (document.datos.lugnro.value==0){
		alert("Debe ingresar el Lugar.");
		document.datos.lugnro.focus();
		return;
	}
	if (document.datos.rubnro.value==0){
		alert("Debe ingresar el Rubro.");
		document.datos.rubnro.focus();
		return;
	}
	*/
<%' End If %>
/*
if ((document.datos.tipmer[0].checked) && (Trim(document.datos.forcal.value) == "")){
	alert("Debe ingresar la Forma de Cálculo.");
	document.datos.forcal.focus();
	}
else if((document.datos.tipmer[0].checked)&&(!stringValido(document.datos.forcal.value))){
	alert("La Forma de Cálculo contiene caracteres inválidos.");
	document.datos.forcal.focus();
	}
else
	{
	*/
	//document.valida.location = "mermas_rubros_con_06.asp?tipo=<%'= l_tipo%>&tipmernro="+document.datos.tipmernro.value + "&lugnro="+document.datos.lugnro.value  + "&rubnro="+document.datos.rubnro.value;
//	}	
valido();
}

function valido(){
	//document.datos.lugnro.disabled = false;
	//document.datos.rubnro.disabled = false;
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.rubnro.focus();
}

function enPesos(){
//Si el tipo de merma es en pesos entonces no puede elegir forma 
	document.datos.forcal.disabled = document.datos.tipmer[1].checked;
	if (document.datos.tipmer[1].checked)
		document.datos.forcal.className = "deshabinp"
	else
		document.datos.forcal.className = "habinp"
}

</script>
<% 

Set l_rs = Server.CreateObject("ADODB.RecordSet")
select Case l_tipo
	Case "A":
		l_rubdes = ""
		l_lugcod = ""
		l_tipmer = "K"
		l_forcal = ""
	Case "M":
		l_tipmernro = request.querystring("cabnro")

		l_sql = "SELECT tkt_tipomerma.rubnro,tkt_tipomerma.lugnro,lugcod,rubdes,tipmer,forcal"
		l_sql = l_sql & " FROM tkt_tipomerma"
		l_sql = l_sql & " INNER JOIN tkt_rubro ON tkt_tipomerma.rubnro = tkt_rubro.rubnro"
		l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_tipomerma.lugnro = tkt_lugar.lugnro"
		l_sql = l_sql & " WHERE tipmernro = " & l_tipmernro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_rubdes = l_rs("rubdes")
			l_lugnro = l_rs("lugnro")
			l_rubnro = l_rs("rubnro")
			l_lugcod = l_rs("lugcod")
			l_forcal = l_rs("forcal")
			l_tipmer = l_rs("tipmer")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="mermas_rubros_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="tipmernro" value="<%= l_tipmernro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Tipos de Mermas para Rubros</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="5%"></td>
				<td width="90%">
					<table cellspacing="0" cellpadding="0" border="0">
					<tr>
					    <td align="right"><b>Lugar:</b></td>
						<td>
							<select <% if l_tipo="M" then%>disabled class="deshabinp"<%end if%> name="lugnro" style="width:300px;">
							<% If l_tipo = "A" then%>
							<option value="">&laquo; Seleccione una opción &raquo;</option>
							<% End If 
							
							l_sql = "SELECT lugnro,lugcod"
							l_sql = l_sql & " FROM tkt_lugar"
							l_sql = l_sql & " ORDER BY lugcod"
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("lugnro")%>"><%=l_rs("lugcod")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							%>
							<script>
								document.datos.lugnro.value = "<%=l_lugnro%>";
							</script>
							</select>
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Rubro:</b></td>
						<td>
							<select <% if l_tipo="M" then%>disabled class="deshabinp"<%end if%> name="rubnro" style="width:300px;">
							<% If l_tipo = "A" then%>
							<option value="">&laquo; Seleccione una opción &raquo;</option>
							<% End If	
										
							l_sql = "SELECT rubnro,rubdes"
							l_sql = l_sql & " FROM tkt_rubro"
							l_sql = l_sql & " ORDER BY rubdes"
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("rubnro")%>"><%=l_rs("rubdes")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							%>
							<script>document.datos.rubnro.value = "<%=l_rubnro%>";</script>
							</select>

						</td>
					</tr>
					<tr>
					    <td nowrap align="right"><b>Tipo de Merma:</b></td>
						<td>
							<input type="Radio" disabled name="tipmer" value="K" <%if UCase(l_tipmer) = "K" then%>checked<%end if%> onclick="javascript:enPesos()">en Kilos
							<input type="Radio" disabled name="tipmer" value="P" <%if UCase(l_tipmer)<> "K" then%>checked<%end if%> onclick="javascript:enPesos()">en Pesos
						</td>
					</tr>					
					<tr>
					    <td nowrap align="right"><b>Forma de Cálculo:</b></td>
						<td>
							<input type="text" disabled class="deshabinp" name="forcal" size="2" maxlength="1" value="<%= l_forcal %>">
						</td>
					</tr>
					<!-- en pesos no puede elegir forma -->

					</table>
				</td>
				<td width="5%"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
		<% call MostrarBoton ("sidebtnABM", "Javascript:Validar_Formulario();","Aceptar")%>
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
