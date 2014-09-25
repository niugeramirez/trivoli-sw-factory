<%@ taglib uri="http://tiles.apache.org/tags-tiles" prefix="tiles" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core"%>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>
<!DOCTYPE html>
<html lang="es" id="app" ng-app="turnosApp">
  <head> 
	<meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    
    <!-- Viewport: Define Ancho, Alto y Escala del Area utilizada por el Navegador para mostrar Contenido incorporando soporte para Dispositivos Moviles-->
    <meta name="viewport" content="width=device-width, initial-scale=1">
    
    <title><spring:message  code="project.title"/></title>
    
	<!-- Estilos del Proyecto -->
    <link href="<c:url value='/resources/css/project.css'/>" rel="stylesheet"/>
    
	<!-- JQuery -->
    <script src="<c:url value='/resources/js/jquery-1.11.0.min.js'/>"></script>
    
	<!-- Bootstrap -->
    <link href="<c:url value='/resources/css/bootstrap.min.css'/>" rel="stylesheet"/>
    <script src="<c:url value='/resources/js/bootstrap.min.js'/>"></script>
     
	<!-- AngularJs -->
    <script src="<c:url value='/resources/js/angular.min-1.2.16.js'/>"></script>

	<!-- JQuery UI -->
	<link href="<c:url value='/resources/css/jqueryui/jquery-ui-1.11.1.min.css'/>" rel="stylesheet"/>
	<script src="<c:url value='/resources/js/jquery-ui-1.11.1.min.js'/>"></script>

	<!-- Js del Proyecto -->
	<script src="<c:url value='/resources/js/project.js'/>"></script>
		
	<!-- HTML5 Shim y Respond.js agregan soporte de IE8 a Elementos HTML5 y Media Queries -->
    <!--[if lt IE 9]>
      <script src="<c:url value='/resources/js/html5shiv.min-3.7.0.js' />"></script>	
      <script src="<c:url value='/resources/js/respond.min-1.4.2.js' />"></script>
	<![endif]-->
  </head>
  
  <body> 
  	<!-- Clase container: Define el Layout principal de todas la Paginas de la Aplicacion -->
	<div class="container">
		<tiles:insertAttribute name="header" />
	    <tiles:insertAttribute name="body" />
		<tiles:insertAttribute name="footer" />
   	</div>
   </body>
</html>