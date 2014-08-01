<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<div class="row-fluid" ng-controller="obrasSocialesController">
probando 
ya agregue el JS
ya corregi tiles xml

		<!-- 		DIV del  diálogo para el AJAX request -->
		<div id="loadingModal" class="modal hide fade in centering" role="dialog" aria-hidden="true">
            <div id="divLoadingIcon" class="text-center">
                <div class="glyphicon glyphicon-align-center loading"></div>
            </div>
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
<!--                     <td class="width15"> -->
<!--                         <div class="text-center"> -->
<!--                             <input type="hidden" value="{{recurso.id}}"/> -->
<!--                             <a href="#editarRecursoDialog" -->
<!--                                ng-click="recursoSeleccionado(recurso);" -->
<!--                                role="button" -->
<%--                                title="<spring:message code="update"/>&nbsp;<spring:message code="recurso"/>" --%>
<!--                                class="btn btn-primary" data-toggle="modal"> -->
<!--                                 <span class="glyphicon glyphicon-pencil"></span> -->
<!--                             </a> -->
<!--                             <a href="#eliminarRecursoDialog" -->
<!--                                ng-click="recursoSeleccionado(recurso);" -->
<!--                                role="button" -->
<%--                                title="<spring:message code="delete"/>&nbsp;<spring:message code="recurso"/>" --%>
<!--                                class="btn btn-primary" data-toggle="modal"> -->
<!--                                 <span class="glyphicon glyphicon-minus"></span> -->
<!--                             </a> -->
<!--                         </div> -->
<!--                     </td> -->
                </tr>
                </tbody>
            </table>
    	    </div>    
		</div>
			
    
</div>
<script src="<c:url value="/resources/js/pages/obrasSociales.js" />"></script>