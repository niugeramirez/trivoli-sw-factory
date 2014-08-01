/**
 * 
 */
package ar.com.trivoli.gestionturnos.obrasocial.controller;

import java.util.Locale;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
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
	@RequestMapping(method = RequestMethod.GET, produces = "application/json")
	public ResponseEntity<?> listAll(@RequestParam int nroPagina, Locale locale) {
		// Se recuperan todas las Obras Sociales
		ListaEntidadDTO<ObraSocial> listaObrasSociales = obraSocialService.recuperarTodos( 
				nroPagina, registrosPorPagina);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<ObraSocial>>(listaObrasSociales,
				HttpStatus.OK);
	}
	
	
	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		return new ModelAndView("admObrasSociales");
	}	
}
