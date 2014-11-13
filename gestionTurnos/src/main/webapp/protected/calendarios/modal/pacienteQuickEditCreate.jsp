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
        		<h3 class="displayInLine" >
            		<div ng-show="modoEditCreate=='edit'"><spring:message code="update"/></div>            		
        			<div ng-show="modoEditCreate=='create'"><spring:message code="create"/></div>            		
        		</h3>        		
    		</div> 
	
			<!--     	DIV contenedor de los parametros de busqueda y los mensajes de error correspondientes -->
    		<div class="modal-body">
        		<form name="pacienteQuickEditCreateForm" novalidate class="well form-horizontal">
					
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
                               	name="dni"
                               	ng-model="pacienteActual.dni"
                               	placeholder="<spring:message code='pacientes.DNI'/>"/>   
                            	<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && pacienteQuickEditCreateForm.dni.$error.required">
                                        <spring:message code="required"/>   
								</span> 
		                 </td>
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	name="apellido"
                               	ng-model="pacienteActual.apellido"
                               	placeholder="<spring:message code='pacientes.apellido'/>"/>  
                            	<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && pacienteQuickEditCreateForm.apellido.$error.required">
                                        <spring:message code="required"/>
                                </span>                               			                 
		                 </td>
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	ng-model="pacienteActual.nombre"
                               	name="nombre"
                               	placeholder="<spring:message code='pacientes.nombre'/>"/>  	
                               	<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && pacienteQuickEditCreateForm.nombre.$error.required">
                                        <spring:message code="required"/>   
								</span> 
		                 </td>
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"                               	
                               	autofocus
                               	ng-model="pacienteActual.nroHistoriaClinica"
                               	placeholder="<spring:message code='pacientes.nroHistoriaClinica'/>"/>  		                 
		                 </td>
		                 <td class="tdRecursosCentered">
		                 	<select class="form-control" 
		                 			ng-model="pacienteActual.obraSocial" 
		                 			ng-options="obraSocial.nombre  for obraSocial in obrasSociales" 
		                 			required 
		                 			name="obraSocial"
		                 			ng-change="mostrarErrorValidacionOS = false"
		                 			>
		                 	</select>
                            <span class="alert alert-danger help-block"
                                     ng-show="mostrarErrorValidacion && mostrarErrorValidacionOS">
                                       <spring:message code="required"/>   
							</span> 		                 	
		                 </td>	
		                 <td class="tdRecursosCentered">
							<input type="text"
						    	class="form-control"
                               	required
                               	autofocus
                               	name="telefono"
                               	ng-model="pacienteActual.telefono"
                               	placeholder="<spring:message code='pacientes.telefono'/>"/>  	
                               	<span class="alert alert-danger help-block"
                                      ng-show="mostrarErrorValidacion && pacienteQuickEditCreateForm.telefono.$error.required">
                                        <spring:message code="required"/>   
								</span>                               		                 
		                 </td>
		             </tr>
		             </tbody>
		         </table>
		 	    </div> 
	            	</div>
            	</form>
            	
        		<span class="alert alert-danger"
		          ng-show="errorSubmit">
					<%-- 		        	<spring:message code="request.error"/> --%>
		        	{{mensajeError}}
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
                   		ng-click="guardarPaciente(pacienteQuickEditCreateForm);"
                   		value='<spring:message code="accept"/>'
                    	/>
            		<button class="btn btn-default"
                    	data-dismiss="modal"
                    	ng-click="exitQuickEditCreate();"
                    	aria-hidden="true">
                		<spring:message code="cancel"/>
            		</button>
    		</div>
    	</div>
    </div>	
</div>    