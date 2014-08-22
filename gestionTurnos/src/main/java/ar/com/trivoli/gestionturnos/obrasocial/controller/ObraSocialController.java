/**
 * 
 */
package ar.com.trivoli.gestionturnos.obrasocial.controller;

import java.util.Locale;

import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

import ar.com.trivoli.gestionturnos.common.controller.ControllerBase;
import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.obrasocial.model.ObraSocial;
import ar.com.trivoli.gestionturnos.obrasocial.service.ObraSocialService;



/**
 * @author ramirez
 *
 *
 *			Controlador de la pagina de obras sociales
 *
 *
 */

@Controller

@RequestMapping(value = "/protected/obrasSociales")
public class ObraSocialController extends ControllerBase<ObraSocial> {

	@Autowired
	private ObraSocialService obraSocialService;

	/**
	 * Método que se invoca al recibir un GET que espera como respuesta un
	 * objeto JSON
	 * 
	 * @param nroPagina
	 *            Numero de Pagina solicitada
	 * @param locale
	 *            Informacion de Localizacion
	 * @return HTTP Response
	 */
	
	/************************************************************************************************************************************************************************/
	private ResponseEntity<?> buscarObrasSociales(	String filtroNombre,
													int nroPagina, 
													Locale locale, 
													String actionMessageKey) {
	
				ListaEntidadDTO<ObraSocial> listaObrasSociales = obraSocialService.buscarObrasSocialesPorNombre(	nroPagina, 
																					registrosPorPagina,
																					filtroNombre);
				
				if (!StringUtils.isEmpty(actionMessageKey)) {
				agregarMensajeAccion(listaObrasSociales, locale, actionMessageKey, null);
				}
				
				Object[] args = { filtroNombre };
				
				agregarMensajeBusqueda(listaObrasSociales, locale,"message.search.for.active", args);
				
				return new ResponseEntity<ListaEntidadDTO<ObraSocial>>(listaObrasSociales,HttpStatus.OK);
	}	
	/************************************************************************************************************************************************************************/	
	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		return new ModelAndView("admObrasSociales");
	}	
	/************************************************************************************************************************************************************************/
	@RequestMapping(method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> listAll(@RequestParam int nroPagina, Locale locale) {
		// Se recuperan todas las Obras Sociales
		ListaEntidadDTO<ObraSocial> listaObrasSociales = obraSocialService.recuperarTodos( 
				nroPagina, registrosPorPagina);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<ObraSocial>>(listaObrasSociales,
				HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/{filtroNombre}", method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> search(
			@PathVariable("filtroNombre") String filtroNombre,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {
		
		return buscarObrasSociales(filtroNombre, nroPagina, locale, null);
	}
	/************************************************************************************************************************************************************************/
	@RequestMapping(value = "/{obraSocialId}", method = RequestMethod.DELETE, produces = "application/json")
	public ResponseEntity<?> delete(
			@PathVariable("obraSocialId") int obraSocialId,
			@RequestParam(required = false) String filtroNombre,
			@RequestParam(required = false, defaultValue = DEFAULT_PAGE_DISPLAYED_TO_USER) int nroPagina,
			Locale locale) {
		
		System.out.println("entra en el delete CON obraSocial id como "+obraSocialId+" filtroNombre "+filtroNombre);
		
		//borro el registro a traves del servicio
		try {
			obraSocialService.delete(obraSocialId);
		} catch (AccessDeniedException e) {
			return new ResponseEntity<Object>(HttpStatus.FORBIDDEN);
		}

		
		//Si hay una búsqueda activa la vuelvo a repetir
		if (existeBusquedaActiva(filtroNombre)) {
			return buscarObrasSociales(filtroNombre, nroPagina, locale,
					"message.delete.success");
		}


		// Se recuperan todos los Recursos
		ListaEntidadDTO<ObraSocial> listaObrasSociales = obraSocialService.recuperarTodos(
				nroPagina, registrosPorPagina);

		agregarMensajeAccion(listaObrasSociales, locale, "message.delete.success",
				null);

		return new ResponseEntity<ListaEntidadDTO<ObraSocial>>(listaObrasSociales,
				HttpStatus.OK);
	}
	/************************************************************************************************************************************************************************/	
}
