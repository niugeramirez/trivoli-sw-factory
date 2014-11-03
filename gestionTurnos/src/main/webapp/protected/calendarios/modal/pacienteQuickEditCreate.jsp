<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<div id = "pacienteQuickEditCreate"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
     
    <!-- 		DIV del  diálogo para el AJAX request -->
	<div id="loadingModalObrasSociales" class="modal hide fade in centering" role="dialog" aria-hidden="true">
           <div id="divLoadingIcon" class="text-center">
               <div class="glyphicon glyphicon-align-center loading"></div>
           </div>
    </div>
    
	<!--     DIV de edicion -->
    <div  class="modal-dialog  modal-lg">	    
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
		                 <th scope="col"><spring:message code="pacientes.telefono"/></th>			                 
		             </tr>
		             </thead>
		             <tbody>
		             <tr>					
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	ng-model="pacienteActual.dni"
                               	placeholder="<spring:message code='pacienteActual.dni'/>"/>   		                 
		                 </td>
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	ng-model="pacienteActual.apellido"
                               	placeholder="<spring:message code='pacienteActual.apellido'/>"/>  		                 
		                 </td>
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	ng-model="pacienteActual.nombre"
                               	placeholder="<spring:message code='pacienteActual.nombre'/>"/>  			                 
		                 </td>
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	ng-model="pacienteActual.nroHistoriaClinica"
                               	placeholder="<spring:message code='pacienteActual.nroHistoriaClinica'/>"/>  		                 
		                 </td>
		                 <td class="tdRecursosCentered">
		                 	<select class="form-control" ng-model="pacienteActual.obraSocial" ng-options="obraSocial.nombre  for obraSocial in obrasSociales"></select>
		                 </td>	
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	ng-model="pacienteActual.telefono"
                               	placeholder="<spring:message code='pacienteActual.telefono'/>"/>  		                 
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
                    	ng-click="exit('#pacienteQuickEditCreate');"
                    	aria-hidden="true">
                		<spring:message code="cancel"/>
            		</button>
    		</div>
    	</div>
    </div>	
</div>    