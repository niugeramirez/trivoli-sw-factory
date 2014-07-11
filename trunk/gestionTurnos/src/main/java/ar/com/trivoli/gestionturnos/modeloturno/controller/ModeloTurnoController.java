package ar.com.trivoli.gestionturnos.modeloturno.controller;

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
import ar.com.trivoli.gestionturnos.modeloturno.model.ModeloTurno;
import ar.com.trivoli.gestionturnos.modeloturno.service.ModeloTurnoService;

/**
 * @author posadas
 * 
 *         Controlador Pagina de Administración de Modelos de Turnos
 */
@Controller
@RequestMapping(value = "/protected/modelos")
public class ModeloTurnoController extends ControllerBase<ModeloTurno> {

	@Autowired
	private ModeloTurnoService modeloTurnoService;

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
		// Se recuperan todos los Modelos
		ListaEntidadDTO<ModeloTurno> listaModelos = modeloTurnoService
				.recuperarTodos(nroPagina, registrosPorPagina);

		// Se arma la Respuesta HTTP
		return new ResponseEntity<ListaEntidadDTO<ModeloTurno>>(listaModelos,
				HttpStatus.OK);
	}

	/**
	 * Método que se invoca al recibir un GET desde el Front
	 * 
	 * @return
	 */
	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		return new ModelAndView("admModelos");
	}
}
