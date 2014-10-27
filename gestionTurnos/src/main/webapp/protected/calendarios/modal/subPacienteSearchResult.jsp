<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

     <div class="modal-dialog  modal-lg modal-content">
	     <div  ng-class="{'alert bg-info': estadoPacientes == 'noresult', 'none': estadoPacientes != 'noresult'}">
	         <h4><span class="glyphicon glyphicon-info-sign"></span> <spring:message code="pacientes.emptyData"/></h4><br/>
	
	         <p><spring:message code="pacientes.emptyData.text"/></p>
	     </div>
    </div>    
    
	<div class="modal-dialog  modal-lg modal-content" ng-class="{'': pacientes.length != '0', 'none': pacientes.length == '0'}">
	
		<!-- 		DIV con la grilla de datos -->           
		<div class="table-responsive" >
		         <table class="table table-bordered table-hover table-condensed">
		             <thead>
		             <tr>
		                 <th scope="col"><spring:message code="pacientes.DNI"/></th>
		                 <th scope="col"><spring:message code="pacientes.apellido"/></th>
		                 <th scope="col"><spring:message code="pacientes.nombre"/></th>
		                 <th scope="col"><spring:message code="pacientes.nroHistoriaClinica"/></th>
		                 <th scope="col"><spring:message code="pacientes.obraSocial"/></th>		                 
		                 <th scope="col"></th>
		             </tr>
		             </thead>
		             <tbody>
		             <tr ng-repeat="registro in pacientes"  title="Asignar Turno a {{registro.apellido}}, {{registro.nombre}}">					
		                 <td class="tdRecursosCentered">{{registro.dni}}</td>
		                 <td class="tdRecursosCentered">{{registro.apellido}}</td>
		                 <td class="tdRecursosCentered">{{registro.nombre}}</td>
		                 <td class="tdRecursosCentered">{{registro.nroHistoriaClinica}}</td>
		                 <td class="tdRecursosCentered">{{registro.obraSocial.nombre}}</td>		                 
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


	<!-- 		DIV de paginado		 -->      
	<div class="text-center">
	      	<button href="#" class="btn btn-primary"
	                  ng-class="{'btn-primary': paginaPacientes.paginaActual != 0, 'disabled': paginaPacientes.paginaActual == 0}"
	                      ng-disabled="paginaPacientes.paginaActual == 0" ng-click="cambiarPaginaPacientes(0)"
	                      title='<spring:message code="pagination.first"/>'
	                      >
	                  <spring:message code="pagination.first"/>
	          </button>
	          <button href="#"
	                      class="btn btn-primary"
	                      ng-class="{'btn-primary': paginaPacientes.paginaActual != 0, 'disabled': paginaPacientes.paginaActual == 0}"
	                      ng-disabled="paginaPacientes.paginaActual == 0" class="btn btn-primary"
	                      ng-click="cambiarPaginaPacientes(paginaPacientes.paginaActual - 1)"
	                      title='<spring:message code="pagination.back"/>'
	                      >&lt;</button>
	          <span>{{paginaPacientes.paginaActual + 1}} <spring:message code="pagination.of"/> {{paginaPacientes.cantPaginas}}</span>
	          <button href="#"
	                      class="btn btn-primary"
	                      ng-class="{'btn-primary': paginaPacientes.cantPaginas - 1 != paginaPacientes.paginaActual, 'disabled': paginaPacientes.cantPaginas - 1 == paginaPacientes.paginaActual}"
	                      ng-click="cambiarPaginaPacientes(paginaPacientes.paginaActual + 1)"
	                      ng-disabled="paginaPacientes.cantPaginas - 1 == paginaPacientes.paginaActual"
	                      title='<spring:message code="pagination.next"/>'
	                      >&gt;</button>
	          <button href="#"
	                      class="btn btn-primary"
	                      ng-class="{'btn-primary': paginaPacientes.cantPaginas - 1 != paginaPacientes.paginaActual, 'disabled': paginaPacientes.cantPaginas - 1 == paginaPacientes.paginaActual}"
	                      ng-disabled="paginaPacientes.cantPaginas - 1 == paginaPacientes.paginaActual"
	                      ng-click="cambiarPaginaPacientes(paginaPacientes.cantPaginas - 1)"
	                      title='<spring:message code="pagination.last"/>'
	                      >
	                  <spring:message code="pagination.last"/>
	          </button>
	        </div>
	</div>