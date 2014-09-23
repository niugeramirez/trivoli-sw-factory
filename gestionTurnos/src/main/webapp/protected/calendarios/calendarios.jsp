<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<div class="row-fluid" ng-controller="controller">



	<!-- 	Titulo y Boton de busqueda -->
	<h2>
        <p class="text-center">
            <spring:message code='turnos.header'/>
        </p>
    </h2>


		<!-- 		DIV del  diálogo para el AJAX request -->
		<div id="loadingModal" class="modal hide fade in centering" role="dialog" aria-hidden="true">
            <div id="divLoadingIcon" class="text-center">
                <div class="glyphicon glyphicon-align-center loading"></div>
            </div>
        </div>
        <!-- 		DIV del  diálogo para el AJAX request -->
		<div id="loadingModalTurno" class="modal hide fade in centering" role="dialog" aria-hidden="true">
            <div id="divLoadingIcon" class="text-center">
                <div class="glyphicon glyphicon-align-center loading"></div>
            </div>
        </div>	
        <!-- 		DIV del  diálogo para el AJAX request -->
		<div id="loadingModalRecursos" class="modal hide fade in centering" role="dialog" aria-hidden="true">
            <div id="divLoadingIcon" class="text-center">
                <div class="glyphicon glyphicon-align-center loading"></div>
            </div>
        </div>	
        
		<!-- 		DIV con algunos mensajes de error como empty data -->
        <div ng-class="{'alert bg-success': mostrarMensajesUsuario == true, 'none': mostrarMensajesUsuario == false}">
        	<h4 class="displayInLine">
                <p class="displayInLine"><span class="glyphicon glyphicon-info-sign"></span>&nbsp;{{pagina.mensajeAccion}}</p>
            </h4>
        </div>

        <div ng-class="{'alert bg-danger': state == 'error', 'none': state != 'error'}">
        	<h4>
            	<span class="glyphicon glyphicon-info-sign"></span> <spring:message code="error.generic.header"/>
            </h4>
            <br/>
            <p><spring:message code="error.generic.text"/></p>
        </div>

        <div ng-class="{'alert bg-info': estado == 'noresult', 'none': estado != 'noresult'}">
            <h4><span class="glyphicon glyphicon-info-sign"></span> <spring:message code="obrasSociales.emptyData"/></h4><br/>

            <p><spring:message code="obrasSociales.emptyData.text"/></p>
        </div>
		<!--         Selector de medicos y fecha -->
 		<form name="filtroRecursoFecha" novalidate class="well form-horizontal">
 			<select class="form-control" ng-model="recursoActual" ng-options="recurso.descripcion for recurso in recursosActuales"></select>
			<input type="date"
		    	class="form-control"
		    	id="txtDescripcion"
               	required
                autofocus
                ng-model="fechaActual"
                name="Fecha"
                data-date-format="dd-mm-yyyy"
                placeholder="Fecha Turnos"/>  			
 		</form>
		

		<!-- 		DIV con la grilla de datos -->           
		<div id="gridContainer" ng-class="{'': estado == 'list', 'none': estado != 'list'}">
			<div class="row show-grid">
			  <div class="col-md-1">
				  <div class="btn-group-vertical" >
					  <div ng-repeat="registroActual in pagina.registros">
					  	<button type="button" class="btn btn-default" ng-click="seleccionarCalendario(registroActual);">{{registroActual.calendario.fechaHoraInicio | date:'hh:mm'}}</button>
					  </div>
				</div>
			  </div>
			  <div class="col-md-4">
					<!-- 		DIV con la grilla de turnos -->           			        
			        <div ng-class="{'alert bg-info': turnosActuales.length == '0', 'none': turnosActuales.length != '0'}">
			            <h4><span class="glyphicon glyphicon-info-sign"></span> <spring:message code="turnos.emptyData"/></h4><br/>
			
			            <p><spring:message code="turnos.emptyData.text"/></p>
			        </div>
			        
					<div id="gridContainer" ng-class="{'': estado == 'list', 'none': estado != 'list'}">
						<div class="table-responsive">
			            <table class="table table-bordered table-hover table-condensed">
			                <tbody>
			                <tr ng-repeat="turno in turnosActuales">
			                    <td class="tdRecursosCentered">{{turno.paciente.apellido}},{{turno.paciente.nombre}} ({{turno.paciente.nroHistoriaClinica}})</td>
			                    <td class="tdRecursosCentered">{{turno.practica.titulo}}</td>
			                    <td class="width15">
			                        <div class="text-center">
										<!--Editar Registros 	------------------------------------------------ -->
			                            <a href="#editarObrasSocialesDialog"
			                               ng-click="obraSocialSeleccionado(obraSocial);"
			                               role="button"
			                               title="<spring:message code="update"/>&nbsp;<spring:message code="obraSocial"/>"
			                               class="btn btn-primary" data-toggle="modal">
			                                <span class="glyphicon glyphicon-pencil"></span>
			                            </a>
			                    </td>
			                    <td class="width15">
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
						<!--     	    Fin del div con la grilla de turnos			   -->
			  </div>
			</div>			
		</div>
	</div>
</div>	
<script src="<c:url value="/resources/js/pages/calendario.js" />"></script>