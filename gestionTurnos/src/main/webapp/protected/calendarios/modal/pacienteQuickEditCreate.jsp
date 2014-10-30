<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

    <div id = "pacienteQuickEditCreate" class="modal-dialog  modal-lg">	    
   		<div class="modal-content"> 
			<div class="modal-header">
        		<h3 class="displayInLine">
            		<spring:message code="search"/>
        		</h3>
    		</div> 
			
			<!--     	DIV contenedor de los parametros de busqueda y los mensajes de error correspondientes -->
    		<div class="modal-body">
        		<form name="busquedaForm" novalidate class="form-horizontal">
					
					<!--         			DIV con la parte de datos del registro actual -->
        			<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion }">
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
		             <tr>					
		                 <td class="tdRecursosCentered">{{registro.dni}}
		                 </td>
		                 <td class="tdRecursosCentered">{{registro.apellido}}
		                 </td>
		                 <td class="tdRecursosCentered">{{registro.nombre}}
		                 </td>
		                 <td class="tdRecursosCentered">{{registro.nroHistoriaClinica}}
		                 </td>
		                 <td class="tdRecursosCentered">{{registro.obraSocial.nombre}}
		                 </td>		                 
		             </tr>
		             </tbody>
		         </table>
		 	    </div> 
	            	</div>
            	</form>
            	
        		<span class="alert alert-danger"
		          ng-show="errorSubmit">
		        	<spring:message code="request.error"/>
			    </span>
	    		<span class="alert alert-danger"
	          		ng-show="errorAccesoIlegal">
	        		<spring:message code="request.illegal.access"/>
	    		</span>
    		</div>
    		
			<!--     		DIV con los botones de aceptar y cancelar -->
    		<div class="modal-footer">
    			<input type="submit"
                   		class="btn btn-primary"
                   		ng-click="buscarPacientes();"
                   		value='<spring:message code="accept"/>'
                    	/>
            		<button class="btn btn-default"
                    	data-dismiss="modal"
                    	ng-click="exit('#pacienteAsign');"
                    	aria-hidden="true">
                		<spring:message code="cancel"/>
            		</button>
    		</div>
    	</div>
    </div>	