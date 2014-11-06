<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<div id="subPacienteSearchResult">    
	<!-- 	DIV con los filtros         -->    
    <div class="modal-dialog  modal-lg modal-content">
	    <div class="col-md-4">
	    	<div ng-class="{'':  mostrarFiltroDNI , 'none': !mostrarFiltroDNI}">
		    <a href="#"
		        role="button"
		        ng-click="resetearBusqueda('dni');"               
		        title="probando"
		        class="btn btn-primary" data-toggle="modal">
		         <span class="glyphicon glyphicon-remove"></span>DNI {{mostrarFiltroDNI}}*
		    </a>
		    </div>
	    </div>

	    <div class="col-md-4">
	    	<div ng-class="{'':  mostrarFiltroApellido , 'none': !mostrarFiltroApellido}">
		    <a href="#"
		        role="button"
		        ng-click="resetearBusqueda('apellido');"               
		        title="probando"
		        class="btn btn-primary" data-toggle="modal">
		         <span class="glyphicon glyphicon-remove"></span>Apellido *{{mostrarFiltroApellido}}*
		    </a>
		    </div>
	    </div>	  

	    <div class="col-md-4">
	    	<div ng-class="{'':  mostrarFiltroNombre , 'none': !mostrarFiltroNombre}">
		    <a href="#"
		        role="button"
		        ng-click="resetearBusqueda('nombre');"               
		        title="probando"
		        class="btn btn-primary" data-toggle="modal">
		         <span class="glyphicon glyphicon-remove"></span>Nombre *{{mostrarFiltroNombre}}*
		    </a>
		    </div>
	    </div>
	    	      
    </div>
    
                
	<!--      DIV de emty data    -->
    <div class="modal-dialog  modal-lg modal-content">
	     <div  ng-class="{'alert bg-info': estadoPacientes == 'noresult', 'none': estadoPacientes != 'noresult'}">
	         <h4><span class="glyphicon glyphicon-info-sign"></span> <spring:message code="pacientes.emptyData"/></h4><br/>
	
	         <p><spring:message code="pacientes.emptyData.text"/></p>
	     </div>
    </div>    
    
    <!-- 		DIV con la grilla de datos -->
	<div class="modal-dialog  modal-lg modal-content" ng-class="{'': pacientes.length != '0', 'none': pacientes.length == '0'}">	
		<h4>
			<p class="text-center">
	        	<spring:message code="message.total.records.found"/>:&nbsp;{{paginaPacientes.totalRegistros}}
	        </p>
        </h4>
        
		<div class="table-responsive" >
		         <table class="table table-bordered table-hover table-condensed">
		             <thead>
		             <tr>
		                 <th scope="col"><spring:message code="pacientes.DNI"/></th>
		                 <th scope="col"><spring:message code="pacientes.apellido"/></th>
		                 <th scope="col"><spring:message code="pacientes.nombre"/></th>
		                 <th scope="col"><spring:message code="pacientes.nroHistoriaClinica"/></th>
		                 <th scope="col"><spring:message code="pacientes.obraSocial"/></th>
		                 <th scope="col"><spring:message code="pacientes.telefono"/></th>			                 
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
		                 <td class="tdRecursosCentered">{{registro.telefono}}</td>			                 
		                 <td class="width15">
		                     <div class="text-center">
								<!--Editar Registros 	------------------------------------------------ -->
		                         <a href="#pacienteQuickEditCreate"
		                            ng-click="quickEditCreatePaciente(registro,'edit');"
		                            role="button"
		                            title="<spring:message code="update"/>&nbsp;<spring:message code="obraSocial"/>"
		                            class="btn btn-primary" data-toggle="modal">
		                             <span class="glyphicon glyphicon-pencil"></span>
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
	</div>