<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<div class="row-fluid" ng-controller="obrasSocialesController">

	<!-- 	Titulo y Boton de busqueda -->
	<h2>
        <p class="text-center">
            <spring:message code='obrasSociales.header'/>
            <a href="#buscarObrasSocialesDialog"
               id="obrasSocialesHeaderButton"
               role="button"
               ng-class="{'': mostrarBotonBuscar == true, 'none': mostrarBotonBuscar == false}"
               title="<spring:message code="search"/>&nbsp;<spring:message code="recurso"/>"
               class="btn btn-primary" data-toggle="modal">
               <span class="glyphicon glyphicon-search"></span> 
            </a>
        </p>
    </h2>
	<!--     Cantidad de registros encontrados -->
    <h4>
        <div ng-class="{'': estado == 'list', 'none': estado != 'list'}">
            <p class="text-center">
                <spring:message code="message.total.records.found"/>:&nbsp;{{pagina.totalRegistros}}
            </p>
        </div>
    </h4>

		<!-- 		DIV del  diálogo para el AJAX request -->
		<div id="loadingModal" class="modal hide fade in centering" role="dialog" aria-hidden="true">
            <div id="divLoadingIcon" class="text-center">
                <div class="glyphicon glyphicon-align-center loading"></div>
            </div>
        </div>
	
		<!-- 	DIV con el mensaje de busqueda         -->
        <div ng-class="{'alert bg-success': mostrarMensajeBusqueda == true, 'none': mostrarMensajeBusqueda == false}">
            <h4>
                <p><span class="glyphicon glyphicon-info-sign"></span>&nbsp;{{pagina.mensajeBusqueda}}</p>
            </h4>
            <a href="#"
               role="button"
               ng-click="resetearBusqueda();"
               ng-class="{'': mostrarMensajeBusqueda == true, 'none': mostrarMensajeBusqueda == false}"
               title="<spring:message code='search.reset'/>"
               class="btn btn-primary" data-toggle="modal">
                <span class="glyphicon glyphicon-remove"></span> <spring:message code="search.reset"/>
            </a>
        </div>
        
		<!-- 		DIV con la grilla de datos -->           
		<div id="gridContainer" ng-class="{'': estado == 'list', 'none': estado != 'list'}">
			<div class="table-responsive">
            <table class="table table-bordered table-hover table-condensed">
                <thead>
                <tr>
                    <th scope="col"><spring:message code="obrasSociales.nombre"/></th>
                    <th scope="col"></th>
                </tr>
                </thead>
                <tbody>
                <tr ng-repeat="obraSocial in pagina.registros">
                    <td class="tdRecursosCentered">{{obraSocial.nombre}}</td>
                    <td class="width15">
                        <div class="text-center">
<!--                             <input type="hidden" value="{{recurso.id}}"/> -->
<!-- 							Editar Registros 	------------------------------------------------ -->
<!--                             <a href="#editarRecursoDialog" -->
<!--                                ng-click="recursoSeleccionado(recurso);" -->
<!--                                role="button" -->
<%--                                title="<spring:message code="update"/>&nbsp;<spring:message code="recurso"/>" --%>
<!--                                class="btn btn-primary" data-toggle="modal"> -->
<!--                                 <span class="glyphicon glyphicon-pencil"></span> -->
<!--                             </a> -->
                            <!--Eliminar Registros 	-------------------------------------------------->
                            <a href="#eliminarObrasSocialesDialog"
                               ng-click="obraSocialSeleccionado(obraSocial);" 
                               role="button"
                               title="<spring:message code="delete"/>&nbsp;<spring:message code="obraSocial"/>"
                               class="btn btn-primary" data-toggle="modal">
                                <span class="glyphicon glyphicon-minus"></span>
                            </a>
                        </div>
                    </td>
                </tr>
                </tbody>
            </table>
    	    </div>    
		</div>

			<!-- 		DIV de paginado		 -->      
			<div class="text-center">
	        	<button href="#" class="btn btn-primary"
	                    ng-class="{'btn-primary': pagina.paginaActual != 0, 'disabled': pagina.paginaActual == 0}"
	                        ng-disabled="pagina.paginaActual == 0" ng-click="cambiarPagina(0)"
	                        title='<spring:message code="pagination.first"/>'
	                        >
	                    <spring:message code="pagination.first"/>
	            </button>
	            <button href="#"
	                        class="btn btn-primary"
	                        ng-class="{'btn-primary': pagina.paginaActual != 0, 'disabled': pagina.paginaActual == 0}"
	                        ng-disabled="pagina.paginaActual == 0" class="btn btn-primary"
	                        ng-click="cambiarPagina(pagina.paginaActual - 1)"
	                        title='<spring:message code="pagination.back"/>'
	                        >&lt;</button>
	            <span>{{pagina.paginaActual + 1}} <spring:message code="pagination.of"/> {{pagina.cantPaginas}}</span>
	            <button href="#"
	                        class="btn btn-primary"
	                        ng-class="{'btn-primary': pagina.cantPaginas - 1 != pagina.paginaActual, 'disabled': pagina.cantPaginas - 1 == pagina.paginaActual}"
	                        ng-click="cambiarPagina(pagina.paginaActual + 1)"
	                        ng-disabled="pagina.cantPaginas - 1 == pagina.paginaActual"
	                        title='<spring:message code="pagination.next"/>'
	                        >&gt;</button>
	            <button href="#"
	                        class="btn btn-primary"
	                        ng-class="{'btn-primary': pagina.cantPaginas - 1 != pagina.paginaActual, 'disabled': pagina.cantPaginas - 1 == pagina.paginaActual}"
	                        ng-disabled="pagina.cantPaginas - 1 == pagina.paginaActual"
	                        ng-click="cambiarPagina(pagina.cantPaginas - 1)"
	                        title='<spring:message code="pagination.last"/>'
	                        >
	                    <spring:message code="pagination.last"/>
	            </button>
	        </div>
		
		<jsp:include page="modal/obrasSocialesDialogs.jsp"/>    
</div>
<script src="<c:url value="/resources/js/pages/obrasSociales.js" />"></script>