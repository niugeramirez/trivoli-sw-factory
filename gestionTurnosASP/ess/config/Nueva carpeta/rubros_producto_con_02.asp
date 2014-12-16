<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: rubros_producto_con_02.asp
'Descripción: Abm de Rubros por producto
'Autor : Gustavo Manfrin
'Fecha: 17/02/2005
'Modificado: 

'Datos del formulario
Dim l_lugnro
Dim l_pronro
Dim l_rubnro
Dim l_bascam
Dim l_tolmax
Dim l_valrefdes
Dim l_valrefhas
Dim l_desporfra
Dim l_concar
Dim l_contra
Dim l_condes
Dim l_oblcar
Dim l_obltra
Dim l_obldes
Dim l_lugcod
Dim l_prodes
Dim l_rubdes

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/ticket/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Rubros por producto - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script>

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_lugnro = request.querystring("cabnro")
l_pronro = request.querystring("pronro")
l_rubnro = request.querystring("rubnro")

l_sql = "SELECT tkt_rub_pro.lugnro, tkt_rub_pro.pronro, tkt_rub_pro.rubnro, bascam, tolmax, valrefdes, valrefhas, desporfra, "
l_sql = l_sql & " concar, contra, condes, oblcar, obldes, obltra,  "
l_sql = l_sql & " tkt_lugar.lugcod, tkt_producto.prodes, tkt_rubro.rubdes "
l_sql = l_sql & " FROM tkt_rub_pro "
l_sql = l_sql & " INNER JOIN tkt_lugar ON tkt_rub_pro.lugnro= tkt_lugar.lugnro "
l_sql = l_sql & " INNER JOIN tkt_producto ON tkt_rub_pro.pronro= tkt_producto.pronro "
l_sql = l_sql & " INNER JOIN tkt_rubro ON tkt_rub_pro.rubnro= tkt_rubro.rubnro "
l_sql = l_sql & " WHERE tkt_rub_pro.lugnro = " & l_lugnro
l_sql = l_sql & " AND tkt_rub_pro.pronro = " & l_pronro
l_sql = l_sql & " AND tkt_rub_pro.rubnro = " & l_rubnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_lugnro=l_rs("lugnro")
    l_rubnro=l_rs("rubnro")
	l_pronro=l_rs("pronro")
	l_bascam=l_rs("bascam")
	l_tolmax=l_rs("tolmax")
	l_valrefdes=l_rs("valrefdes")
	l_valrefhas=l_rs("valrefhas")
	l_desporfra=l_rs("desporfra")
	l_concar = l_rs("concar")
	l_contra = l_rs("contra")
	l_condes = l_rs("condes")
	l_oblcar = l_rs("oblcar")
	l_obltra = l_rs("obltra")
	l_obldes = l_rs("obldes")
	l_lugcod=l_rs("lugcod")
	l_prodes=l_rs("prodes")
	l_rubdes=l_rs("rubdes")
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
	    <td class="th2"  nowrap>Rubros por producto</td>
		<td class="th2" align="right">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td colspan="2" height="100%">
			<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
				<tr>
					<td width="50%"></td>
					<td>
						<table>
							<tr>
							    <td align="right" nowrap><b>Lugar:</b></td>
								<td>
									<input type="text" readonly class="deshabinp" name="lugcod" size="20" maxlength="20" value="<%= l_lugcod %>">
								</td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>Producto:</b></td>
								<td>
									<input type="text" readonly class="deshabinp" name="prodes" size="50" maxlength="50" value="<%= l_prodes %>">
								</td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>Rubro:</b></td>
								<td>
									<input type="text" readonly class="deshabinp" name="rubdes" size="20" maxlength="20" value="<%= l_rubdes%>">
								</td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>Valor de la base:</b></td>
								<td>
									<input type="text" readonly class="deshabinp" name="valbase" size="20" maxlength="20" style="text-align : right;" value="<%= l_bascam %>">
								</td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>Máximo valor sin merma:</b></td>
								<td>
									<input type="text" readonly class="deshabinp" name="tolmax" size="20" maxlength="20" style="text-align : right;" value="<%= l_tolmax %>">
								</td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>Rango de valores:</b></td>
								<td>
									<input type="text" readonly class="deshabinp" name="refdes" size="20" maxlength="20" style="text-align : right;" value="<%= l_valrefdes %>">
									&nbsp;
									<input type="text" readonly class="deshabinp" name="refhas" size="20" maxlength="20" style="text-align : right;" value="<%= l_valrefhas %>">
								</td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>Descuento por fracción:</b></td>
								<td>
									<input type="text" readonly class="deshabinp" name="despor" size="20" maxlength="20" style="text-align : right;" value="<%= l_desporfra %>">
								</td>
							</tr>		
							<tr>
								<td colspan="2">
									<table style="border: thin solid Silver;">
										<tr>
										    <td align="center">&nbsp;</td>
										    <td align="center"><b>Considera</b></td>
										    <td align="center"><b>Obligatorio</b></td>
										</tr>
										<tr>
										    <td align="right"><b>Carga:</b></td>
										    <td align="center"> <input type="Checkbox" disabled name="ccarga" <% If l_concar = -1 then  %>checked<% end if %>></td>
										    <td align="center"> <input type="Checkbox" disabled name="vcarga" <% If l_oblcar= -1 then  %>checked<% end if %>></td>
										</tr>
										<tr>
										    <td align="right"><b>Descarga:</b></td>
										    <td align="center"> <input type="Checkbox" disabled name="cdesc" <% If l_condes = -1 then  %>checked<% end if %>></td>
										    <td align="center"> <input type="Checkbox" disabled name="vdesc" <% If l_obldes= -1 then  %>checked<% end if %>></td>
										</tr>
										<tr>
										    <td align="right"><b>Tránsito:</b></td>
										    <td align="center"> <input type="Checkbox" disabled name="ctran" <% If l_contra = -1 then  %>checked<% end if %>></td>
										    <td align="center"> <input type="Checkbox" disabled name="vtran" <% If l_obltra= -1 then  %>checked<% end if %>></td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
					<td width="50%"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
	    <td colspan="2" align="right" class="th2">
			<a class=sidebtnABM href="Javascript:window.close()">Salir</a>
		</td>
	</tr>
</table>
</form>
<%
set l_rs = nothing
Cn.Close
Cn = nothing
%>
</body>
</html>
