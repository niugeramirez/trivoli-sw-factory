<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/ay_confrep.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/util.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/numero.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : rep_recibo_liq_laslennas.asp
Descripcion    : Modulo que se encarga de generar los recibos de sueldo en formato A4
Creador        : Scarpa D.
Fecha Creacion : 16/04/2004
Modificacion   : Martín Ferraro - Cambios de Categoria por OS elegida, label basico por convenido
	28-06-05 - Fernando Favre - Se achico la letra en Categoria y OSocial.
	05/04/2006 - Martin Ferraro	- Se agrego funcion para imprimir por rango
-----------------------------------------------------------------------------
-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<style>
.stytit00
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 10pt;
	FONT-WEIGHT: bold;
	text-align: right;
}
.stydat01-l
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: left;	
	padding-left: 5px; 	
}
.stydat01-c
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: center;	
}
.stydat01-r
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: right;	
	padding-right: 5px; 
}
.stydat01-l-ll
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: left;	
	padding-left: 5px; 	
    border-left-style: solid;
	border-left-width: 1px;
	border-left-color: Black;	
}
.stydat01-c-ll
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: center;	
    border-left-style: solid;
	border-left-width: 1px;
	border-left-color: Black;	

}
.stydat01-c-tl
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: center;	
    border-top-style: solid;
	border-top-width: 1px;
	border-top-color: Black;	

}
.stydat01-r-ll
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: right;	
	padding-right: 5px; 
    border-left-style: solid;
	border-left-width: 1px;
	border-left-color: Black;	
}
.stydat01-r-tll
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: right;	
	padding-right: 5px; 
    border-left-style: solid;
	border-left-width: 1px;
	border-left-color: Black;	
	
    border-top-style: solid;
	border-top-width: 1px;
	border-top-color: Black;	
	
}
.stydat01-l-tll
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: normal;
    text-align: left;	
	padding-right: 5px; 
    border-left-style: solid;
	border-left-width: 1px;
	border-left-color: Black;	
	
    border-top-style: solid;
	border-top-width: 1px;
	border-top-color: Black;	
	
}
.stytit01-l
{
	background-color: #333399;
	COLOR: #ffffff;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: bold;
	text-align: left;
}
.stytit01-c
{
	background-color: #CCCCCC;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: bold;
	text-align: center;
}
.stydat02-l
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 7pt;
	FONT-WEIGHT: normal;
    text-align: left;	
}
.stydat02-c
{
	background-color: #FFFFFF;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 7pt;
	FONT-WEIGHT: normal;
    text-align: center;	
}
.stytit02-l
{
	background-color: #333399;
	COLOR: #ffffff;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 7pt;
	FONT-WEIGHT: bold;
	text-align: left;
}
.stytit02-c
{
	background-color: #CCCCCC;
	COLOR: #000000;
	FONT-FAMILY: "Arial";
	FONT-SIZE: 8pt;
	FONT-WEIGHT: bold;
	text-align: center
}

</style>


<%
on error goto 0

function Mes(nro)
   select case nro
     case 1 mes = "Enero"
     case 2 mes = "Febrero"
     case 3 mes = "Marzo"
     case 4 mes = "Abril"
     case 5 mes = "Mayo"
     case 6 mes = "Junio"
     case 7 mes = "Julio"
     case 8 mes = "Agosto"
     case 9 mes = "Septiembre"
     case 10 mes = "Octubre"
     case 11 mes = "Noviembre"
     case 12 mes = "Diciembre"
   end select
end function 'mes(nro)

function valor(dato)
  if isNull(dato) then
     valor = ""
  else
     valor = dato
  end if
end function 'valor(dato)

Dim l_cont
    l_cont = 0

'Variables Locales
Dim l_i
Dim l_bpronro
Dim l_sql
Dim l_rs
Dim l_rs2
Dim l_rs3
Dim l_rs10
Dim l_cant_conceptos 
Dim l_max_conceptos 
Dim l_actual_concepto

Dim l_total_remun
Dim l_total_noremun
Dim l_total_desc
Dim l_total_unidad

Dim l_total_remun_actual
Dim l_total_noremun_actual
Dim l_total_desc_actual
Dim l_total_unidad_actual

Dim l_ancho_recibo

'Tamano para Carta
'l_ancho_recibo = 450

'Tamano para A4
l_ancho_recibo = 485

Dim l_emplegfijo
Dim l_cant_leye

l_bpronro = request("bpronro")
l_emplegfijo  = request("empleg")

'Variables que se usan en el recibo
Dim l_apellido
Dim l_Nombre
Dim l_direccion
Dim l_Legajo
Dim l_pliqnro
Dim l_pliqmes
Dim l_pliqanio
Dim l_pliqdepant
Dim l_pliqfecdep
Dim l_pliqbco
Dim l_cuil
Dim l_empfecalta
Dim l_sueldo
Dim l_categoria
Dim l_centrocosto
Dim l_localidad
Dim l_profecpago
Dim l_formapago
Dim l_ternro
Dim l_pronro
Dim l_cliqnro
Dim l_textoFinRecibo
Dim l_empnombre 
Dim l_empdire 
Dim l_empcuit
Dim l_emplogo 
Dim l_emplogoalto 
Dim l_emplogoancho
Dim l_empfirma 
Dim l_empfirmaalto 
Dim l_empfirmaancho
Dim l_empfecbaja
Dim l_regimenhor
Dim l_obra_social
Dim l_calificacion
Dim l_reportaa
Dim l_lugarpago
Dim l_sector
Dim l_subcategoria

Dim l_mostrar_firma

'-------------------------------------------------------------------------------------------------------------------
'INICIO:
'Descripcion: imprime un recibo de sueldo
'-------------------------------------------------------------------------------------------------------------------
sub imprimirRecibo
%>

<table width="100%">
<tr>
  <td>
    <!-- Datos Empresa -->
	<table>
	  <tr>
	    <td align="left">
  		 <img src="<%= l_emplogo%>" height="<%= l_emplogoalto%>" width="<%= l_emplogoancho%>">
		</td>
	  </tr>	
	</table>    
  </td>
  <td class="stydat01-l" style="font-size:7pt;">
 	  <%= l_empnombre%><br>
	  <%= l_empdire%><br>
	  <%= l_empcuit%>  
  </td>
  <td valign="top">
     <table border="0" cellpadding="1" width="100%"  cellspacing="1" style="border-color:gray ; border-width: 1 ; border-style:solid">
	   <tr>
		 <td class="stytit01-c" nowrap>
		 Lugar y Fecha de Pago
		 </td>  
	   </tr>
		<tr>  
		  <td class="stydat01-c">
			   <%= l_lugarpago & "&nbsp;" &  l_profecpago %>
		  </td>		 
		</tr>
	  </table>
  </td>  
</tr>
</table> 
<!-- 
------------------------------------------------------------------------------------------------------------------------
  -->
<table width="100%">   
  <td>
   <!-- Datos Empleado -->
    <table border="0" cellpadding="1" width="100%"  cellspacing="1" style="border-color:gray ; border-width: 1 ; border-style:solid">
		<tr>
		  <td class="stytit01-c">
		  Legajo
		  </td>		  
		  <td class="stytit01-c" nowrap>
		  Apellido y Nombres
		  </td>
		  <td class="stytit01-c">
		  Documento
		  </td>		  
		  <td class="stytit01-c">
		  Sector
		  </td>		  
		</tr>
		<tr>
		  <td class="stydat01-c" style="font-size:9pt;">
		  <b><%= l_legajo%></b>
		  </td>				  
		  <td class="stydat01-l" nowrap>
		  <%= l_apellido & ", " & l_nombre %>
		  </td>
		  <td class="stydat01-c" nowrap>
		  <b>CUIL:&nbsp;</b><%= l_cuil%>
		  </td>
		  <td class="stydat01-c">
		  <%= l_sector%>
		  </td>
		</tr>
	</table>  
  </td>
</tr>
</table> 
<!-- 
------------------------------------------------------------------------------------------------------------------------
  -->
<table width="100%">
<tr>
  <td align="right">

	<table width="100%" border="0" cellpadding="1"  cellspacing="1" align="left" style="height:42px;border-color:gray ; border-width: 1 ; border-style:solid">
	<tr>
	 <td class="stytit01-c" width="30%" nowrap>
	   Per&iacute;odo Abonado
	 </td>  
	 <td class="stytit02-c" nowrap>
	   &Uacute;ltimo Dep&oacute;sito Cargas Sociales
	 </td>  
	</tr>  
	<tr>
	 <td class="stydat01-c">
 	   <%= Mes(l_pliqmes) & "&nbsp;&nbsp;" & l_pliqanio %>&nbsp;
	 </td>  
	 <td>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" style="height:21px">
		<tr>
		 <td class="stytit02-c" style="height:10px">
		   Fecha
		 </td>  
		 <td class="stytit02-c">
		   Per&iacute;odo
		 </td>  
		 <td class="stytit02-c">
		   Banco
		 </td>  	 
		</tr>  
		<tr>
		 <td class="stydat02-c">
		   <%= l_pliqfecdep%>
		 </td>  
		 <td class="stydat02-c">
		   <%= l_pliqdepant%>
		 </td>  
		 <td class="stydat02-c">
		   <%= l_pliqbco%>
		 </td>  	 
		</tr>  	
		</table>
		 </td>  
	</tr>  
	</table>
  </td>
</tr>
</table>
<!-- 
------------------------------------------------------------------------------------------------------------------------
  -->
<table width="100%">   
  <td>
   <!-- Datos Empleado -->
    <table border="0" cellpadding="1" width="100%"  cellspacing="1" style="border-color:gray ; border-width: 1 ; border-style:solid">
		<tr>
		  <td class="stytit01-c" >
		  Categor&iacute;a
		  </td>		  
		  <td class="stytit01-c">
		  Sub-Categor&iacute;a
		  </td>
		  <td class="stytit01-c" width="25%">
		  OS Elegida
		  </td>
		  <td class="stytit01-c">
		  Básico
		  </td>		  
		  <td class="stytit01-c">
		  Fecha de Ing.
		  </td>		  
		</tr>
		<tr>
		  <td class="stydat02-c" nowrap>
		  <%= UCASE(l_calificacion)%>
		  </td>				  
		  <td class="stydat02-c" nowrap>
		  <%= l_subcategoria %>
		  </td>
		  <td class="stydat02-c" nowrap>
		  <%= l_obra_social  %>
		  </td>
		  <td class="stydat01-c" nowrap>
		  <% if not isnull(l_sueldo) then response.write formatnumber(l_sueldo,2) end if%>
		  </td>
		  <td class="stydat01-c" nowrap>
		  <%= l_empfecalta%>
		  </td>
		</tr>
	</table>  
  </td>
</tr>
</table> 
<!-- 
------------------------------------------------------------------------------------------------------------------------
  -->
<table width="100%">
<tr>
  <td width="100%">
	<table width="100%" border="0" cellpadding="0"  cellspacing="0" style="width:<%= l_ancho_recibo%>px;border-color:gray ; border-width: 1 ; border-style:solid">
	<tr>
	 <td class="stytit01-c" width="1%" style="padding-left:10px;">
	   C&oacute;digo
	 </td>
	 <td class="stytit01-c">
	   Concepto
	 </td>
	 <td class="stytit01-c">
	   Cantidad
	 </td>
	 <td class="stytit01-c">
	   Suj.&nbsp;a&nbsp;Ret.
	 </td>
	 <td class="stytit01-c">
	   Rem.&nbsp;Exen.
	 </td>
	 <td class="stytit01-c">
	   Descuentos
	 </td>	 
    </tr>
	<% 

	   do until l_rs2.eof OR (l_cant_conceptos = l_max_conceptos)
          l_cant_conceptos = l_cant_conceptos + 1
		  
	%>
		<tr>
		 <td class="stydat01-r">
		   <%= l_rs2("conccod")%>
		 </td>
		 <td class="stydat01-l-ll">
		   <%= l_rs2("concabr")%>
		 </td>
		 <td class="stydat01-r-ll">
			 <% if FormatNumber(CDbl(l_rs2("dlicant")),2) > 0 then%>
				   <%= FormatNumber(CDbl(l_rs2("dlicant")),2) %>
			 <% else%>
				   &nbsp;
			 <% end if%>
		   <% l_total_unidad = l_total_unidad + CDbl(l_rs2("dlicant")) %> 	   
		 </td>
		 <% select case CInt(l_rs2("conctipo")) %>
		    <% case 0 %>
				 <td class="stydat01-r-ll">&nbsp;
			      
  				 </td>
				 <td class="stydat01-r-ll">&nbsp;
			      
				 </td>
				 <td class="stydat01-r-ll">&nbsp;
			      
				 </td>			 			
			
			<% case 1 ' Remunerativo%>
				 <td class="stydat01-r-ll">
				   <%= FormatNumber(CDbl(l_rs2("dlimonto")),2)%>
				   <% l_total_remun = l_total_remun + CDbl(l_rs2("dlimonto")) %> 
				 </td>
				 <td class="stydat01-r-ll">&nbsp;
			      
				 </td>
				 <td class="stydat01-r-ll">&nbsp;
			      
				 </td>			 
			
			<% case 2 ' No Remunerativo%>			
				 <td class="stydat01-r-ll">&nbsp;
			      
				 </td>
				 <td class="stydat01-r-ll">
				   <%= FormatNumber(CDbl(l_rs2("dlimonto")),2)%>
				   <% l_total_noremun = l_total_noremun + CDbl(l_rs2("dlimonto")) %> 
				 </td>
				 <td class="stydat01-r-ll">&nbsp;				 
			      
				 </td>			 
			
			<% case 3 ' Descuento%>			
				 <td class="stydat01-r-ll">&nbsp;
			      
				 </td>
				 <td class="stydat01-r-ll">&nbsp;				 
			      
				 </td>			 
				 <td class="stydat01-r-ll">
				   <%= FormatNumber(CDbl(l_rs2("dlimonto")),2)%>
				   <% l_total_desc = l_total_desc + CDbl(l_rs2("dlimonto")) %> 
				 </td>
		 
	     <% end select %>

	    </tr>
	<% 
	  l_rs2.moveNext
	loop
	
	if l_rs2.eof then
		'Lleno con blancos el resto del recibo
		   for l_i = l_cant_conceptos to l_max_conceptos 
%>		   
		<tr>
		 <td class="stydat01-c">
	       <br>
		 </td>
		 <td class="stydat01-l-ll">
	       <br>
		 </td>
		 <td class="stydat01-r-ll">
	       <br>
		 </td>
		 <td class="stydat01-r-ll">
	       <br>	   
		 </td>
		 <td class="stydat01-r-ll">
	       <br>
		 </td>
		 <td class="stydat01-r-ll">
	       <br>
		 </td>
	    </tr>
	    <% next %>
	<%end if
	  'Si no es fin de archivo hay transporte
	%>
		 <tr>
		 <td class="stydat01-l-tll" colspan="6">
	       	<%if l_rs2.eof then
		 	l_sql = " SELECT * FROM rep_recibo_leyenda "
			l_sql = l_sql & " WHERE bpronro = " & l_bpronro
			l_sql = l_sql & "   AND  pronro = " & l_pronro
		 	l_sql = l_sql & "   AND  ternro = " & l_ternro
		   
			rsOpen l_rs3, cn, l_sql, 0
		   
		 	do until l_rs3.eof
		   		 response.write  "&nbsp;" 
		   		 response.write l_rs3("leyenda") & "<br>" 
		 		l_rs3.movenext
			loop
		   
		 	l_rs3.close
			
	   		end if %>
		 </td>
		</tr>	
		<tr>
		 <td class="stydat01-c-tl" colspan="3">
	       <b>Totales</b>
		 </td>
		 <td class="stydat01-r-tll">
	       <%= FormatNumber(l_total_remun,2)%>
		 </td>
		 <td class="stydat01-r-tll">
	       <%= FormatNumber(l_total_noremun,2)%>
		 </td>
		 <td class="stydat01-r-tll">
	       <%= FormatNumber(l_total_desc,2)%>
		 </td>		 
	    </tr>	
	  </table>
    </td>
  </tr>
  <!--
  *************************************************************
  -->  
  <tr>
    <td colspan="5">
	  <table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-color:gray ; border-width: 1 ; border-style:solid">
	    <tr>
		  <td valign="top" class="stydat01-l" style="background:#f3f3f3;">

			<%if l_rs2.eof then%>
		       <b>Importe Neto Pagado en Letras</b><%= l_textoFinRecibo%><br>
			   <% if ( l_total_remun + l_total_noremun + l_total_desc ) < 0 then%>
			   <% response.write "Menos " & NumerosALetras(FormatNumber(( l_total_remun + l_total_noremun + l_total_desc ) * (-1))) %><br>	   
			   <% else%>
			   <% response.write  NumerosALetras(FormatNumber( ( l_total_remun + l_total_noremun + l_total_desc  ) )) %><br>
			   <% end if%>
			<%end if%>
			<br>
			<%= l_formapago %>
		  </td>  
		  <td valign="top" class="stydat01-c-ll" width="10%">
		    <table cellpadding="0" border="0">
			<tr>
			 <td class="stydat01-c" nowrap>
			   <b>Neto&nbsp;Pagado</b>
			 </td>
		    </tr>
			<tr>
			 <td class="stydat01-c" nowrap>
			   <b>
			   <%if l_rs2.eof then
			       response.write FormatNumber(( l_total_remun + l_total_noremun + l_total_desc ),2)
			     end if %>
				</b>
			 </td>
		    </tr>	
		    </table>
		  </td>
		</tr>
	  </table>	
	</td>
  </tr>    
  <!--
  *************************************************************
  -->  
  <tr>
    <td colspan="5">
	  <table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-color:gray ; border-width: 1 ; border-style:solid">
	    <tr>
		  <td class="stydat01-l" nowrap valign="top" width="80%">
			    <% if l_mostrar_firma then %>
				     Por&nbsp;<%= l_empnombre %>
				<%else%>
				     Recib&iacute; el pago de la presente<br>
					 liquidaci&oacute;n y duplicado de este recibo
			    <%end if%>			  
		  </td>		  
		  <td class="stydat01-c" nowrap style="padding-top:5px;">
 		        &nbsp;
			    <% if l_mostrar_firma then %>
				       <%if l_empfirma = "" then  %>
					        <br><br><br>
					   <%else%>
			                <img src="<%= l_empfirma%>" height="<%= l_empfirmaalto%>" width="<%= l_empfirmaancho%>"><br>
				       <%end if%>
				      <hr width="100" style="heigth:3px">
				      firma del empleador
				<%else%>
				       <%if l_empfirma = "" then  %>
					        <br><br><br>
					   <%else%>
			                <b style="height:<%= l_empfirmaalto%>px;width:<%= l_empfirmaancho%>px;">
							<br>&nbsp;
							</b>					   
				       <%end if%>
				      <hr width="100" style="heigth:3px">
				      firma del empleado				
			    <%end if%>	
				&nbsp;		  
		  </td>		
		  <td class="stydat01-c" nowrap valign="bottom">
		        &nbsp;&nbsp;
			    <% if l_mostrar_firma then %>
				   <b>Duplicado</b>
				<%else%>
				   <b>Original</b>				
			    <%end if%>	
				&nbsp;&nbsp;		  
		  </td>				  
		</tr>
	   <td class="stydat01-l" align="left">
       Recibo N&deg;&nbsp;<%= l_cliqnro%>
  	   </td>
	  </table>
	</td>
  </td>
  </tr>    
  
</table>	
<!-- 
------------------------------------------------------------------------------------------------------------------------
  -->
<table width="100%">
</table>

<%
end sub 'imprimirRecibo
'-------------------------------------------------------------------------------------------------------------------
'FIN:
'Descripcion: imprime un recibo de sueldo
'-------------------------------------------------------------------------------------------------------------------

%>
<head>
	<title>Recibo de sueldo</title>
</head>

<body>

<%
'-------------------------------------------------------------------------------------------------------------------
' EMPIEZA
'-------------------------------------------------------------------------------------------------------------------
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs10 = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")

'Busco los datos de los recibos
    l_sql = " SELECT * FROM rep_recibo WHERE bpronro = " & l_bpronro 
    if l_emplegfijo <> "" then
       l_sql = l_sql & " AND legajo = " & l_emplegfijo
	end if
    l_sql = l_sql & " ORDER BY orden "		

rsOpen l_rs, cn, l_sql, 0

do until l_rs.eof

    l_ternro      = l_rs("ternro")
    l_pronro      = l_rs("pronro")	
	l_apellido    = l_rs("apellido")
	l_Nombre      = l_rs("nombre")
	l_direccion   = l_rs("direccion")
	l_Legajo      = l_rs("legajo")
	l_pliqnro     = l_rs("pliqnro")
	l_pliqmes     = l_rs("pliqmes")
	l_pliqanio    = l_rs("pliqanio")
	l_pliqdepant  = l_rs("pliqdepant")
	l_pliqfecdep  = l_rs("pliqfecdep")
	l_pliqbco     = l_rs("pliqbco")
	l_cuil        = l_rs("cuil")
	l_empfecalta  = l_rs("empfecalta")
	l_sueldo      = l_rs("sueldo")
	l_categoria   = l_rs("categoria")
	l_centrocosto = l_rs("centrocosto")
	l_localidad   = l_rs("localidad")
	l_profecpago  = l_rs("profecpago")
	l_formapago   = l_rs("formapago")
	l_empnombre   = l_rs("empnombre")
	l_empdire     = l_rs("empdire")
	l_empcuit     = l_rs("empcuit")
	l_emplogo     = l_rs("emplogo")
	l_emplogoalto = l_rs("emplogoalto")
	l_emplogoancho= l_rs("emplogoancho")
	l_empfirma    = l_rs("empfirma")
	l_empfirmaalto = l_rs("empfirmaalto")
	l_empfirmaancho= l_rs("empfirmaancho")
	l_calificacion = l_rs("categoria")
	l_obra_social   = valor(l_rs("auxchar1"))
	l_regimenhor   = valor(l_rs("auxchar2"))
	l_lugarpago    = valor(l_rs("auxchar3"))
	l_sector    = valor(l_rs("auxchar4"))

	
	'Busco la cantidad de leyendas del recibo
    l_sql = " SELECT * FROM rep_recibo_leyenda "
    l_sql = l_sql & " WHERE bpronro = " & l_bpronro
    l_sql = l_sql & "   AND  pronro = " & l_pronro
    l_sql = l_sql & "   AND  ternro = " & l_ternro
   
    rsOpen l_rs3, cn, l_sql, 0
   
    l_cant_leye = 0
    do until l_rs3.eof
      l_cant_leye = l_cant_leye + 1
      l_rs3.movenext
    loop
    l_rs3.close
	
	'Busco Sub-Categoria del empleado

    l_sql = " SELECT * FROM his_estructura "
    l_sql = l_sql & " INNER JOIN estructura ON  estructura.estrnro = his_estructura.estrnro "
    l_sql = l_sql & " WHERE  ternro = " & l_ternro
    l_sql = l_sql & " AND his_estructura.htetdesde <=" & cambiafecha(l_profecpago,"YMD",true) & " AND "
    l_sql = l_sql & " (his_estructura.htethasta >= " & cambiafecha(l_profecpago,"YMD",true) & " OR his_estructura.htethasta IS NULL) "
    l_sql = l_sql & " AND his_estructura.tenro  = 44 "
    rsOpen l_rs10, cn, l_sql, 0   
    do until l_rs10.eof
      l_subcategoria = l_rs10("estrdabr") 
      l_rs10.movenext
    loop	
    l_rs10.close


	'Busco los conceptos del empleado en el proceso

    'l_sql = " SELECT * ,CAST (conccod as int) orden FROM rep_recibo_det "
    'l_sql = l_sql & " WHERE bpronro = " & l_bpronro
	'l_sql = l_sql & "   AND pronro  = " & l_pronro
	'l_sql = l_sql & "   AND ternro  = " & l_ternro
	'l_sql = l_sql & " ORDER BY orden "
	
    l_sql = " SELECT * FROM rep_recibo_det "
    l_sql = l_sql & " WHERE bpronro = " & l_bpronro
	l_sql = l_sql & "   AND pronro  = " & l_pronro
	l_sql = l_sql & "   AND ternro  = " & l_ternro
	
	rsOpen l_rs2, cn, l_sql, 0
	
    l_cant_conceptos = 0
    l_max_conceptos  = 0
	l_actual_concepto = 0
    l_total_remun = 0
    l_total_noremun = 0
    l_total_desc = 0	
	l_total_unidad = 0
    l_total_remun_actual = 0
    l_total_noremun_actual = 0
    l_total_desc_actual = 0	
    l_total_unidad_actual = 0
	
	do until l_rs2.eof

	   l_cliqnro = l_rs2("cliqnro")

	   'Indico cual es la cantidad maxima que se pueden imprimir en el recibo
	   l_max_conceptos = l_max_conceptos + 26

%>
<table style="border-color:white ; border-width: 2 ; border-style:solid">

<tr>
<td width="49%" style="width:<%= (l_ancho_recibo + 5)%>px;border-color:black ; border-width: 2px ; border-style:solid">
<!-- EMPIEZA - Recibo Empleado -->
<% 

l_textoFinRecibo = ""
l_cant_conceptos = l_actual_concepto
l_rs2.moveFirst
l_rs2.move l_actual_concepto
l_total_remun  = l_total_remun_actual
l_total_noremun  = l_total_noremun_actual
l_total_desc  = l_total_desc_actual
l_total_unidad = l_total_unidad_actual
l_mostrar_firma = false
imprimirRecibo 

%>
<!-- TERMINA - Recibo Empleado -->
</td>
<td width="2%" nowrap>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td>
<td width="49%" style="width:<%= (l_ancho_recibo + 5)%>px;border-color:black ; border-width: 2px ; border-style:solid ; margin-top: 5px;">
<!-- EMPIEZA - Recibo Empresa -->
<% 

l_textoFinRecibo = ""
l_cant_conceptos = l_actual_concepto
l_rs2.moveFirst
l_rs2.move l_actual_concepto
l_total_remun  = l_total_remun_actual
l_total_noremun  = l_total_noremun_actual
l_total_desc  = l_total_desc_actual
l_total_unidad = l_total_unidad_actual
l_mostrar_firma = true
imprimirRecibo 

'Actualizo en que concepto me encuentro para el proximo recibo del empleado, por si existe transporte
l_actual_concepto = l_cant_conceptos
l_total_remun_actual  = l_total_remun
l_total_noremun_actual  = l_total_noremun
l_total_desc_actual  = l_total_desc
l_total_unidad_actual = l_total_unidad

%>
<!-- TERMINA - Recibo Empresa -->
</td>
</tr>

</table>

<%

      if not l_rs2.eof then 
	     response.write "<p style='page-break-before:always'></p>"
	  end if
   loop
   
   l_rs2.close
   
   l_rs.MoveNext
   
   if not l_rs.eof then 
	     response.write "<p style='page-break-before:always'></p>"
   end if

loop

l_rs.close

%>
<div id="objetos" name="objetos">

</div>

<script>
var indice=1;

function ImprYa(){
    var WebBrowser = '<OBJECT ID="WebBrowser' + indice + '" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
    document.all.objetos.insertAdjacentHTML('beforeEnd', WebBrowser);

    execScript("on error resume next: WebBrowser" + indice + ".ExecWB 6, 2", "VBScript");
	indice++;	
	
	document.all.objetos.innerHTML = '';	
}
  parent.ifrmListo=1;
</script>

</body>
</html>
