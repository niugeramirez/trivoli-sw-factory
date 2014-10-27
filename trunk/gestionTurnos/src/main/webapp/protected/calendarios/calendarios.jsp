<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<div class="row-fluid" ng-controller="controller">

	<!-- 	Titulo -->
	<h2>
        <p class="text-center">
            <spring:message code='turnos.header'/>
        </p>
    </h2>



		<!-- 		DIV del  diálogo para el AJAX request -->
		<div id="loadingModalCalendario" class="modal hide fade in centering" role="dialog" aria-hidden="true">
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
        
		<!-- 		DIV con algunos mensajes de error  -->
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


		<!--         Selector de medicos y fecha -->
 		<div class="col-md-12">
 			<div class="col-lg-5"> 
 				<select class="form-control" ng-model="recursoActual" ng-options="recurso.descripcion for recurso in recursosActuales" ng-change="listarCalendario()"></select>
			</div>
			<div class="col-lg-3"> 
				<input type="text" datepicker ng-model="fechaActual" ng-change="listarCalendario()"/>
			</div>
 		</div>

		<!-- 		DIV con la grilla de datos -->     
		<div ng-class="{'alert bg-info': calendariosActuales.length == '0', 'none': calendariosActuales.length != '0'}">
            <h4><span class="glyphicon glyphicon-info-sign"></span> <spring:message code="turnos.calendario.emptyData"/></h4><br/>

            <p><spring:message code="turnos.calendario.emptyData.text"/></p>
        </div>      
		<div id="gridContainer" ng-class="{'': calendariosActuales.length != '0', 'none': calendariosActuales.length == '0'}">
			<div class="row show-grid">
				<div class="col-md-1">		
					<div ng-repeat="registroActual in calendariosActuales" class="btn-group-vertical">			  	
						<a href="javascript:void(0)"  class="btn btn-default"  ng-click="seleccionarCalendario(registroActual);"  ng-class="{ 'active' : calendarioActual.id == registroActual.calendario.id }">{{registroActual.calendario.fechaHoraInicio | date:'hh:mm'}}</a>				 
					</div>
			  	</div>
		  
			  <div class="col-md-4">
			  		<!--     	Botn de creacion -->    	
			    	<div ng-class="text-center">
			            <br/>
			            <a href="#pacienteAsign"
			               role="button"
			               title="<spring:message code='create'/>&nbsp;<spring:message code='turno'/>"
			               class="btn btn-primary"
			               data-toggle="modal">
			               <span class="glyphicon glyphicon-plus"></span>
			               &nbsp;&nbsp;<spring:message code="create"/>&nbsp;<spring:message code="turno"/>
			            </a>
			        </div>
					<!-- 		DIV con la grilla de turnos -->           			        
			        <div ng-class="{'alert bg-info': turnosActuales.length == '0', 'none': turnosActuales.length != '0'}">
			            <h4><span class="glyphicon glyphicon-info-sign"></span> <spring:message code="turnos.emptyData"/></h4><br/>
			
			            <p><spring:message code="turnos.emptyData.text"/></p>
			        </div>
			        
					<div id="gridContainer" ng-class="{'': turnosActuales.length != '0', 'none': turnosActuales.length == '0'}">
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
	<jsp:include page="modal/pacienteAsign.jsp"/>
</div>	
<script src="<c:url value="/resources/js/pages/calendario.js" />"></script>