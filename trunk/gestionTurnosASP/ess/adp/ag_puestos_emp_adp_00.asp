<% Option Explicit %>
<%	'<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->  %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/numero.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-------------------------------------------------------------------------------------------
Archivo        : ag_puestos_emp_adp_00.asp
Descripcion    : Muestra el puesto del empleado con el link asociado a la descripcion PDF
Autor          : Gustavo Ring
Fecha          : 26/07/2007
Modificación   :
--------------------------------------------------------------------------------------------
-->
<% 
on error goto 0

' Variables
Dim l_ternro
dim l_rs
dim l_sql
Dim l_empleg
Dim l_habilitar_estado

l_habilitar_estado = (Session("empleg") <> l_ess_empleg)

' Filtro
  Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
  Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
  Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)

' Orden
  Dim l_Orden      ' Son las etiquetas que aparecen en el orden
  Dim l_CamposOr   ' Son los campos para el orden
  
' Filtro
  l_etiquetas = "Licencia:;Apellido:;Fecha desde:;Fecha hasta:"
  l_Campos    = "tipdia.tddesc;empleado.terape;elfechadesde;elfechahasta"
  l_Tipos     = "T;T;F;F"

' Orden
  l_Orden     = "Licencia:;Apellido:;Fecha desde:;Fecha hasta:"
  l_CamposOr  = "tipdia.tddesc;empleado.terape;elfechadesde;elfechahasta"
  
  l_empleg = request("empleg")
%>

<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<title>Puestos - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
      <table border="0" cellpadding="0" cellspacing="0" height="95%">
        <tr>
          <th colspan="2" align="left">Puestos</th>
        </tr>
        <tr valign="top" height="100%">
          <td colspan="4" style="">
      	  <iframe name="ifrm" src="ag_puestos_emp_adp_01.asp?empleg=<%= l_empleg%>" width="100%" height="100%"></iframe> 		  
	      </td>
        </tr>
        <tr valign="top">
          <td colspan="4"  height="20">
	      </td>
        </tr>

      </table>
</body>
</html>
