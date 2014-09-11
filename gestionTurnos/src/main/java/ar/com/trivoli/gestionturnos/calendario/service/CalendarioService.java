/**
 * 
 */
package ar.com.trivoli.gestionturnos.calendario.service;

import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.calendario.model.CalendarioDTO;
import ar.com.trivoli.gestionturnos.calendario.repository.ICalendarioRepository;
import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.turno.service.TurnoService;


/**
 * @author ramirez
 *
 */

@Service
@Transactional
public class CalendarioService {

	@Autowired
	private ICalendarioRepository calendarioRepository;
	
	@Autowired
	private TurnoService turnoService;
	
	/************************************************************************************************************************************************************************/
//	private Sort ordenPredeterminado() {
//		return new Sort(Sort.Direction.ASC, "fechaHoraInicio");
//	}	
	/************************************************************************************************************************************************************************/
	@Transactional(readOnly = true)
	public ListaEntidadDTO<CalendarioDTO> recuperarTodos() {
				

		//traigo los calendarios del repositorio
		List<Calendario> resultado = (List<Calendario>) calendarioRepository.findAll();
		
		//recorro los calendarios y armo la lista de calendarios DTO con sus turnos	
		List<CalendarioDTO> resultadoDTO = new ArrayList<CalendarioDTO>(); 
				
		for (Calendario cal :resultado) {
			
			resultadoDTO.add(new CalendarioDTO(cal, turnoService.buscarTurnosPorCalendario(cal)));
		}

		
		return new ListaEntidadDTO<CalendarioDTO>(	1
													,resultadoDTO.size()
													,resultadoDTO);
	}	
	/************************************************************************************************************************************************************************/
}
