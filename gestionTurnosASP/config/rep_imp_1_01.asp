<% Option Explicit
'if request.querystring("excel") then
'	Response.AddHeader "Content-Disposition", "attachment;filename=Estadisticas.xls" 
'	Response.ContentType = "application/vnd.ms-excel"
'end if
 %>

<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 53
Const l_cantcols = 10
Const l_empresa = "Cámara Portuaria y Marítima de Bahía Blanca"


Dim l_porcentaje
Dim l_rs
Dim l_rs2
Dim l_buqdes
Dim l_canbuq
Dim l_totton
Dim l_cantidad_buques

Dim l_tipobuque
Dim l_cantipbuq

Dim l_sql
dim primero
dim ultimo


Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte 

dim l_total 

dim l_fecini
dim l_fecfin

'Variable usadas para imprimir los Totales
dim l_nroope
dim l_anioini

' Imprime los Totales


'Obtengo los parametros
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")

'l_repelegido  = request.querystring("repnro")

l_anioini = "01/01/" & year(l_fecfin)


Dim l_indice_exportadora

Dim l_fila

Function NombreTipoOperacion(nro)

select case nro
case 1
	NombreTipoOperacion = "Carga"
case 2
	NombreTipoOperacion = "Descarga"
case 3
	NombreTipoOperacion = "Exportación"
case 4
	NombreTipoOperacion = "Importación"
end select

end Function

Function NombreMes(nro)

select case nro
case 1
	NombreMes = "Ene"
case 2
	NombreMes = "Feb"
case 3
	NombreMes = "Mar"
case 4
	NombreMes = "Abr"
case 5
	NombreMes = "May"
case 6
	NombreMes = "Jun"
case 7
	NombreMes = "Jul"
case 8
	NombreMes = "Ago"
case 9
	NombreMes = "Sep"
case 10
	NombreMes = "Oct"
case 11
	NombreMes = "Nov"
case 12
	NombreMes = "Dic"

end select

end Function


sub Inicializar_Arreglo(Arr, Lim, Valor)

	for x = 1 to Lim
		Arr(x) = Valor
	next

end sub	



sub encabezado_expbuq(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" >
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td align="left" width="100%" colspan="7">
							<b>* <%= titulo%></b> 
						</td>
				       	<td align="right" nowrap width="5%" > 
							<!--P&aacute;gina: <%'= l_nropagina%> -->
						</td>				
					</tr>
					<!--
					<tr>
						<td align="left" width="100%" colspan="7">
							<%'= l_fecini  %>&nbsp;-&nbsp;<%'= l_fecfin %>
						</td>
				       	<td align="right" nowrap width="5%" > 
							&nbsp;
						</td>										
					</tr>
					<tr>
				       	<td nowrap colspan="8">&nbsp;
						</td>				
					</tr>
					-->														
				</table>
			</td>				
		</tr>
<%
end sub 'encabezado

sub fin_encabezado
%>
</table>	
<%
end sub 'finencabezado


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
function Sitio(sitnro){
	
	param = "qfecini=<%= l_fecini %>&qfecfin=<%= l_fecfin %>&qsitnro=" + sitnro ;
	
   	abrirVentana("rep_imp_2_00.asp?" + param ,'',780,580);	
	//parent.frames.ifrm.focus();
	//window.print();	
}

function Destino(desnro){
	
	param = "qfecini=<%= l_fecini %>&qfecfin=<%= l_fecfin %>&qdesnro=" + desnro ;
	
   	abrirVentana("rep_imp_3_00.asp?" + param ,'',780,580);	
	//parent.frames.ifrm.focus();
	//window.print();	
}

function Mercaderia(mernro){
	
	param = "qfecini=<%= l_fecini %>&qfecfin=<%= l_fecfin %>&qmernro=" + mernro ;
	
   	abrirVentana("rep_imp_4_00.asp?" + param ,'',780,580);	
	//parent.frames.ifrm.focus();
	//window.print();	
}


function Cant_Buques(){
	
	param = "qfecini=<%= l_fecini %>&qfecfin=<%= l_fecfin %>";
	
   	abrirVentana("rep_imp_detalle_buques_00.asp?" + param,'',700,550);	
	//parent.frames.ifrm.focus();
	//window.print();	
}

function Exportadora(expnro){1
	
	param = "qfecini=<%= l_fecini %>&qfecfin=<%= l_fecfin %>&qexpnro=" + expnro ;
	
   	abrirVentana("rep_imp_5_00.asp?" + param ,'',780,580);	
	//parent.frames.ifrm.focus();
	//window.print();	
}

function Agencia(agenro){1
	
	param = "qfecini=<%= l_fecini %>&qfecfin=<%= l_fecfin %>&qagenro=" + agenro ;
	
   	abrirVentana("rep_imp_6_00.asp?" + param ,'',780,580);	
	//parent.frames.ifrm.focus();
	//window.print();	
}


</script>
	
	
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

encabezado_expbuq("TOTAL DE BUQUES") 

l_sql =  " SELECT count(distinct(buqdes)) "
l_sql = l_sql & " FROM buq_buque  "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " AND buq_buque.tipopenro = 4 "
rsOpen l_rs, cn, l_sql, 0
if not l_Rs.eof then
	l_cantidad_buques = l_rs(0)
else
	l_cantidad_buques = 0
end if 

%>
<tr>
	<td align="center" width="100%"><a href="Javascript:Cant_Buques(<%= l_rs(0) %>);"><img alt="Ver Detalle de Buques" src="/serviciolocal/shared/images/cal.gif" border="0"></a>&nbsp;&nbsp;&nbsp;<b><%= l_cantidad_buques %></b></td>
</tr>	
<%

fin_encabezado
l_rs.close

encabezado_expbuq("TIPOS DE BUQUES") 

'l_encabezado = true
'l_corte = false
'l_total = 0


l_sql = " SELECT DISTINCT buq_buque.buqdes, buq_tipobuque.tipbuqdes "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " INNER JOIN buq_tipobuque ON buq_tipobuque.tipbuqnro = buq_buque.tipbuqnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
l_sql = l_sql & " ORDER BY buq_tipobuque.tipbuqdes "
'response.write l_sql

rsOpen l_rs, cn, l_sql, 0

l_canbuq = 0
l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Tipo Buque</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Cantidad</th>
    </tr>
<%	
if not l_rs.eof then 
	l_tipobuque = l_rs(1)
end if

l_cantipbuq = 0
do until l_rs.eof
	if l_tipobuque <> l_rs(1) then
%>
	<tr>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_tipobuque %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_cantipbuq %></td>
    </tr>
<%
		l_tipobuque = l_rs(1)
		l_cantipbuq = 0
	end if
	l_cantipbuq = l_cantipbuq  + 1
	
'	l_nrolinea = l_nrolinea + 1
'	l_canbuq = l_canbuq + 1
'	l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
%>
	<tr>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_tipobuque %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_cantipbuq %></td>
    </tr>
<!--
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%'= l_canbuq %></b></td>
</tr>
-->
<%	
fin_encabezado
l_rs.close

encabezado_expbuq("PRODUCTO") 

l_sql = " SELECT buq_mercaderia.merdes , sum(conton), buq_mercaderia.mernro "
l_sql = l_sql & " FROM buq_buque  "
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_mercaderia ON buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
l_sql = l_sql & " group by buq_mercaderia.merdes, buq_mercaderia.mernro "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Producto</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="left" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><a href="Javascript:Mercaderia(<%= l_rs(2) %>);"><img alt="Ver Detalle del Producto" src="/serviciolocal/shared/images/cal.gif" border="0"></a>&nbsp;&nbsp;&nbsp;<%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close

encabezado_expbuq("SITIO") 

l_sql = " SELECT buq_sitio.sitdes , sum(conton), buq_sitio.sitnro "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_sitio ON buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
l_sql = l_sql & " group by buq_sitio.sitdes , buq_sitio.sitnro "
l_sql = l_sql & " order by 2 desc "
rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Sitio</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
		
	<tr>
		<td align="left" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><a href="Javascript:Sitio(<%= l_rs(2) %>);"><img alt="Ver Detalle del Sitio" src="/serviciolocal/shared/images/cal.gif" border="0"></a>&nbsp;&nbsp;&nbsp;<%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%

fin_encabezado
l_rs.close

encabezado_expbuq("DESTINO") 

l_sql = " SELECT buq_destino.desdes , sum(conton), buq_destino.desnro "
l_sql = l_sql & " FROM buq_buque " 
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_destino ON buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
l_sql = l_sql & " group by buq_destino.desdes , buq_destino.desnro "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Sitio</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="left" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><a href="Javascript:Destino(<%= l_rs(2) %>);"><img alt="Ver Detalle del Destino" src="/serviciolocal/shared/images/cal.gif" border="0"></a>&nbsp;&nbsp;&nbsp;<%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close

encabezado_expbuq("EXPORTADORA") 

l_sql = " SELECT buq_exportadora.expdes , sum(conton), buq_exportadora.expnro "
l_sql = l_sql & " FROM buq_buque  "
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_exportadora ON buq_exportadora.expnro = buq_contenido.expnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
l_sql = l_sql & " group by buq_exportadora.expdes , buq_exportadora.expnro "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Exportadora</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="left" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><a href="Javascript:Exportadora(<%= l_rs(2) %>);"><img alt="Ver Detalle Exportadora" src="/serviciolocal/shared/images/cal.gif" border="0"></a>&nbsp;&nbsp;&nbsp;<%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close

encabezado_expbuq("AGENCIA") 

l_sql = " SELECT buq_agencia.agedes , sum(conton), buq_agencia.agenro "
l_sql = l_sql & " FROM buq_buque  "
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_agencia ON buq_agencia.agenro = buq_buque.agenro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
l_sql = l_sql & " group by buq_agencia.agedes , buq_agencia.agenro "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Exportadora</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="left" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><a href="Javascript:Agencia(<%= l_rs(2) %>);"><img alt="Ver Detalle Agencia" src="/serviciolocal/shared/images/cal.gif" border="0"></a>&nbsp;&nbsp;&nbsp;<%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close

response.end



l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Cantidad de Buques</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Total Toneladas</td>				
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><b><%= l_totton %></b></td>
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;">&nbsp;</td>			
</tr>
<%
l_nrolinea = l_nrolinea + 1
response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

