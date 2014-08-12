<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<div id="editarModeloDialog"
     class="modal fade"
     role="dialog"
     aria-hidden="true">
	 <div class="modal-dialog modal-lg">
   		<div class="modal-content">     
		    <div class="modal-header">
		        <h3 id="editarModeloModalLabel" class="displayInLine">
		            <spring:message code="details"/>&nbsp;<spring:message code="modelo"/>
		        </h3>
		    </div>
		    <div class="modal-body">
		        <form name="editarModeloForm" novalidate class="well form-horizontal">
		            <input type="hidden"
		                   required
		                   ng-model="modelo.id"
		                   name="id"
		                   value="{{modelo.id}}"/>
					
<!-- 					<div class="form-group" ng-class="{'has-error': mostrarErrorValidacion && editarModeloForm.descripcion.$error.required}"> -->
<%-- 						<label for="txtDescripcion" class="col-lg-3 control-label"><spring:message code="modelos.descripcion"/>:</label> --%>
<!-- 						<div class="col-lg-7"> -->
<!-- 							<input type="text" -->
<!-- 						    	class="form-control" -->
<!-- 						    	id="txtDescripcion" -->
<!--                                	required -->
<!--                                	autofocus -->
<!--                                	ng-model="modelo.descripcion" -->
<!--                                	name="descripcion" -->
<%--                                	placeholder="<spring:message code='modelos.descripcion'/>"/>          --%>
<!-- 						</div> -->
<!-- 						<div class="col-lg-2"> -->
<!-- 							<label> -->
<!--                             	<span class="alert alert-danger help-block" -->
<!--                                       ng-show="mostrarErrorValidacion && editarModeloForm.descripcion.$error.required"> -->
<%--                                         <spring:message code="required"/> --%>
<!--                                 </span> -->
<!--                         	</label> -->
<!-- 						</div> -->
<!-- 					</div> -->
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
			<div class="modal-footer">
			 <input type="submit"
		     	class="btn btn-primary"
		        ng-click="updateModelo(editarModeloForm);"
		        value='<spring:message code="update"/>'/>
		     <button class="btn btn-default"
		     	data-dismiss="modal"
		        ng-click="exit('#editarModeloDialog');"
		        aria-hidden="true">
		        <spring:message code="cancel"/>
		     </button>
		</div>
		</div>
	</div>				    
</div>