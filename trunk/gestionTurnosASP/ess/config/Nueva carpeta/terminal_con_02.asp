<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 
'Archivo: terminal_con_02.asp
'Descripción: ABM de Terminales
'Autor : Gustavo Manfrin
'Fecha: 13/04/2005
'Modificado: 
'alert(document.datos.teritkt2[document.datos.teritkt2.selectedIndex].text);

'on error goto 0


'Datos del formulario
Dim l_ternro
Dim l_terdes
Dim l_tercod
Dim l_tercov
Dim l_tersect
Dim l_terbalf1
Dim	l_balanza1
Dim l_tercomf1
Dim l_tercomfcf1
Dim l_tervvcf1
Dim l_terbalf2
Dim	l_balanza2
Dim l_tercomf2
Dim l_tercomfcf2
Dim l_tervvcf2
Dim l_terimptick
Dim l_terimpcpor
Dim l_terimpremi
Dim l_terimpetiq
Dim l_terresmacp
Dim l_planro
Dim l_terizq


Dim l_terimptktnro
Dim l_terimpcpornro
Dim l_terimpremnro
Dim l_terimpetinro


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
<title><%= Session("Titulo")%>Terminales - Ticket</title>
</head>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
if (Trim(document.datos.tercod.value) == ""){
	alert("Debe ingresar el Código.");
	document.datos.tercod.focus();
	}
else if(!stringValido(document.datos.tercod.value)){
	alert("El Código contiene caracteres inválidos.");
	document.datos.tercod.focus();
	}
else if(Trim(document.datos.terdes.value) == ""){
	alert("Debe ingresar la Descripción.");
	document.datos.terdes.focus();
	}
else if(!stringValido(document.datos.terdes.value)){
	alert("La Descripción contiene caracteres inválidos.");
	document.datos.terdes.focus();
	}
else if(Trim(document.datos.planro.value) == ""){
	alert("Debe ingresar una Planta.");
	document.datos.planro.focus();
	}
else{
	var d=document.datos;
	document.valida.location = "terminal_con_06.asp?tipo=<%= l_tipo%>&ternro="+document.datos.ternro.value + "&tercod="+document.datos.tercod.value + "&terdes="+document.datos.terdes.value;
	}	
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.tercod.focus();
}

function Imp_ticket(){
	document.datos.teritkt.value = document.datos.teritkt2[document.datos.teritkt2.selectedIndex].text;
}

function Imp_CartaPorte(){
	document.datos.tericpor.value = document.datos.tericpor2[document.datos.tericpor2.selectedIndex].text;
}

function Imp_Remito(){
	document.datos.teriremi.value = document.datos.teriremi2[document.datos.teriremi2.selectedIndex].text;
}

function Imp_Etiqueta(){
	document.datos.terieti.value = document.datos.terieti2[document.datos.terieti2.selectedIndex].text;
}
 

</script>
<% 
select Case l_tipo
	Case "A":
		l_ternro = ""
		l_terdes = ""
		l_tercod = ""
		l_tercov = ""
		l_tersect = ""
		l_terbalf1 = ""
		l_balanza1 = ""		
		l_tercomf1 = ""
		l_tercomfcf1 = ""
		l_tervvcf1 = ""
		l_terbalf2 = ""
    	l_balanza2 = ""		
		l_tercomf2 = ""
		l_tercomfcf2 = ""
		l_tervvcf2 = ""
		l_terimptick = ""
		l_terimpcpor = ""
		l_terimpremi = ""
		l_terimpetiq = ""
		l_terresmacp = ""
		l_planro = ""
		l_terizq = "2"
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_ternro = request.querystring("cabnro")
		
		l_sql = "SELECT  tkt_planta.plades, tkt_terminal.* , tkt_balanza.balcod as balanza1, aliasbal.balcod as balanza2" ', terizq"
		l_sql = l_sql & " FROM tkt_terminal "
		l_sql = l_sql & " LEFT JOIN tkt_planta ON tkt_terminal.planro= tkt_planta.planro "
  	    l_sql = l_sql & " LEFT JOIN tkt_balanza ON tkt_terminal.terbalf1= tkt_balanza.balnro "
 	    l_sql = l_sql & " LEFT JOIN tkt_balanza aliasbal ON tkt_terminal.terbalf2= aliasbal.balnro "
		l_sql  = l_sql  & " WHERE ternro = " & l_ternro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_terdes = l_rs("terdesC")
			l_tercod = l_rs("tercod")
			l_tercov = l_rs("tercov")
			l_tersect = l_rs("tersect")
			l_terbalf1 = l_rs("terbalf1")
			l_tercomf1 = l_rs("tercomf1")
			l_balanza1 = l_rs("balanza1")			
			l_tercomfcf1 = l_rs("tercomfcf1")
			l_tervvcf1 = l_rs("tervvcf1")
			l_terbalf2 = l_rs("terbalf2")
			l_balanza2 = l_rs("balanza2")						
			l_tercomf2 = l_rs("tercomf2")
			l_tercomfcf2 = l_rs("tercomfcf2")
			l_tervvcf2 = l_rs("tervvcf2")
			l_terimptick = trim(l_rs("terimptick"))
			l_terimpcpor = trim(l_rs("terimpcpor"))
			l_terimpremi = trim(l_rs("terimpremi"))
			l_terimpetiq = trim(l_rs("terimpetiq"))
			l_terresmacp = l_rs("terresmacp")
			l_planro = l_rs("planro")
'			if isnull(l_rs("terizq"))  then 
'				l_terizq = "2"
'			else 
'				l_terizq = l_rs("terizq")
'			end if
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.tercod.focus()">
<form name="datos" action="terminal_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="ternro" value="<%= l_ternro %>">

<input type="Hidden" name="teritkt" value="<%= l_terimptick %>">
<input type="Hidden" name="tericpor" value="<%= l_terimpcpor %>">
<input type="Hidden" name="teriremi" value="<%= l_terimpremi %>">
<input type="Hidden" name="terieti" value="<%= l_terimpetiq %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" nowrap>Terminales</td>
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
							<input type="text" name="tercod" size="18" maxlength="15" value="<%= l_tercod %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Descripción:</b></td>
						<td>
							<input type="text" name="terdes" size="60" maxlength="50" value="<%= l_terdes %>">
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Transporte:</b></td>
						<td>
							<input type="Radio" name="tercov" value="C" <%if l_tercov = "C" then%>checked<%end if%> >Camión
							<input type="Radio" name="tercov" value="V" <%if l_tercov = "V" then%>checked<%end if%> >Vagón
						</td>
					</tr>
					<tr>
					    <td align="right"><b>Sector:</b></td>
						<td>
							<input type="text" name="tersec" size="2" maxlength="1" value="<%= l_tersect %>">				
						</td>
					</tr>
					<tr>
  					  <td colspan="2">
					  <table style="border: thin solid Silver;">
					  					
						<tr>
						<td>
						<!--<b>Izq:<b><input type="Radio" name="izq" value="2" <%if l_terizq = "2" then%>checked<%end if%>>-->
						</td>
						
					    <td align="right" nowrap><b>Balanza F2:</b></td>
						<td>
						<select name="terbal2" style="width:150px;">
                            <option value="">&nbsp;</option>							
							<% 
							l_sql = "SELECT balnro,balcod"
							l_sql = l_sql & " FROM tkt_balanza"
							l_sql = l_sql & " ORDER BY balnro"
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("balnro")%>"><%=l_rs("balcod")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							%>
							<script>
								document.datos.terbal2.value = "<%=l_terbalf2%>";
							</script>
							</select>
                        </td>
					    <td align="right" ><b>Puerto:</b></td>
						<td>
							<input type="text" name="tercom2" size="2" maxlength="1"  value="<%= l_tercomf2 %>">				
						</td>

					    <td align="right" nowrap><b>Puerto FC:</b></td>
						<td>
							<input type="text" name="tercomfc2"  size="2" maxlength="1"  value="<%= l_tercomfcf2 %>">				
						</td>

					    <td align="right" nowrap><b>Vuelta a cero: </b></td>
						 <td align="center"> 
						    <input type="Checkbox" readonly disabled name="tervvc2" <% If l_tervvcf2 = -1 then  %>checked<% end if %>>
						</td>
						
					</tr>	
						<tr>						
						<td>
						<!--<b>Izq:<b><input type="Radio" name="izq" value="1" <%if l_terizq = "1" then%>checked<%end if%>>-->
						</td>						
					    <td align="right" nowrap><b>Balanza F1:</b></td>
						<td>
						<select name="terbal1" style="width:150px;">
                            <option value="">&nbsp;</option>							
							<% 
							l_sql = "SELECT balnro,balcod"
							l_sql = l_sql & " FROM tkt_balanza"
							l_sql = l_sql & " ORDER BY balnro"
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("balnro")%>"><%=l_rs("balcod")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							%>
							<script>
								document.datos.terbal1.value = "<%=l_terbalf1%>";
							</script>
							</select>
                        </td>
					    <td align="right" ><b>Puerto:</b></td>
						<td>
							<input type="text" name="tercom1" size="2" maxlength="1"  value="<%= l_tercomf1 %>">				
						</td>

					    <td align="right" nowrap><b>Puerto FC:</b></td>
						<td>
							<input type="text" name="tercomfc1"  size="2" maxlength="1"  value="<%= l_tercomfcf1 %>">				
						</td>

					    <td align="right" nowrap><b>Vuelta a cero: </b></td>
						 <td align="center"> 
						    <input type="Checkbox" readonly disabled name="tervvc1" <% If l_tervvcf1 = -1 then  %>checked<% end if %>>
						</td>
					</tr>	
                    </table> 						
					</tr>
					<tr>
	  			    <td colspan="2">
			 	    <table>
					<tr>
					    <td align="right" nowrap><b>Impresora Ticket:</b></td>
							<td>
				  		   <select name="teritkt2" style="width:390px;" onchange="Javascript:Imp_ticket();" >
							<% If l_tipo = "A" then
							Set l_rs = Server.CreateObject("ADODB.RecordSet")%>
							<option value="">&laquo; Seleccione una opción &raquo;</option>
							<% End If 
							
							l_sql = "SELECT impnro,impnom"
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " ORDER BY impnom "
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option  value="<%=l_rs("impnro")%>"><%=l_rs("impnom")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							
							' Obtengo el codigo de la Impresora segun su nombre
							l_sql = "SELECT impnro "
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " WHERE impnom = '" & l_terimptick & "'"
							rsOpen l_rs, cn, l_sql, 0
							l_terimptktnro = 0
							if not l_rs.eof then
								l_terimptktnro = l_rs("impnro")
							end if
							l_rs.close
							%>
							<script>
								document.datos.teritkt2.value = "<%= l_terimptktnro %>";
							</script>
							</select>
                        </td>				
					</tr>				
					<tr>
					    <td align="right" nowrap><b>Impresora C.Porte:</b></td>
							<td>
				  		   <select name="tericpor2" style="width:390px;" onchange="Javascript:Imp_CartaPorte();">
							<% If l_tipo = "A" then
							Set l_rs = Server.CreateObject("ADODB.RecordSet")%>
							<option value="">&laquo; Seleccione una opción &raquo;</option>
							<% End If 
							
							l_sql = "SELECT impnro,impnom"
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " ORDER BY impnom "
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("impnro")%>"><%=l_rs("impnom")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							
							' Obtengo el codigo de la Impresora segun su nombre
							l_sql = "SELECT impnro "
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " WHERE impnom = '" & l_terimpcpor & "'"
							rsOpen l_rs, cn, l_sql, 0
							l_terimpcpornro = 0
							if not l_rs.eof then
								l_terimpcpornro = l_rs("impnro")
							end if
							l_rs.close
							
							%>
							<script>
								document.datos.tericpor2.value = "<%= l_terimpcpornro %>";
							</script>
							</select>
                        </td>
					    <td align="right"><b>Resma:</b></td>
						<td>
							<input type="text" name="terrcpor" size="2" maxlength="1" value="<%= l_terresmacp %>">
						</td>
					</tr>
					<tr>
					    <td align="right" nowrap><b>Impresora Remito:</b></td>
							<td>
				  		   <select name="teriremi2" style="width:390px;" onchange="Javascript:Imp_Remito();">
							<% If l_tipo = "A" then
							Set l_rs = Server.CreateObject("ADODB.RecordSet")%>
							<option value="">&laquo; Seleccione una opción &raquo;</option>
							<% End If 
							
							l_sql = "SELECT impnro,impnom"
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " ORDER BY impnom "
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("impnro")%>"><%=l_rs("impnom")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							
							' Obtengo el codigo de la Impresora segun su nombre
							l_sql = "SELECT impnro "
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " WHERE impnom = '" & l_terimpremi & "'"
							rsOpen l_rs, cn, l_sql, 0
							l_terimpremnro = 0
							if not l_rs.eof then
								l_terimpremnro = l_rs("impnro")
							end if
							l_rs.close							
							%>
							<script>
								document.datos.teriremi2.value = "<%=l_terimpremnro %>";
							</script>
							</select>
                        </td>				
					</tr>
					<tr>
					    <td align="right" nowrap><b>Impresora Etiqueta:</b></td>
							<td>
				  		   <select name="terieti2" style="width:390px;" onchange="Javascript:Imp_Etiqueta();">
							<% If l_tipo = "A" then
							Set l_rs = Server.CreateObject("ADODB.RecordSet")%>
							<option value="">&laquo; Seleccione una opción &raquo;</option>
							<% End If 
							
							l_sql = "SELECT impnro,impnom"
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " ORDER BY impnom "
							rsOpen l_rs, cn, l_sql, 0 
							do while not l_rs.eof
							%>
								<option value="<%=l_rs("impnro")%>"><%=l_rs("impnom")%></option>
							<%
								l_rs.MoveNext
							loop
							l_rs.close
							
							' Obtengo el codigo de la Impresora segun su nombre
							l_sql = "SELECT impnro "
							l_sql = l_sql & " FROM tkt_impresora "
							l_sql = l_sql & " WHERE impnom = '" & l_terimpetiq & "'"
							rsOpen l_rs, cn, l_sql, 0
							l_terimpetinro = 0
							if not l_rs.eof then
								l_terimpetinro = l_rs("impnro")
							end if
							l_rs.close
							
							%>
							<script>
								document.datos.terieti2.value = "<%= l_terimpetinro %>";
							</script>
							</select>
                        </td>				
					</tr>					
										
                	</table>									
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
