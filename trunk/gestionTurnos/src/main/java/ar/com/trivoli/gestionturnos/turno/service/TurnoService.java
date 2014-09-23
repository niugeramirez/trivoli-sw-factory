/**
 * 
 */
package ar.com.trivoli.gestionturnos.turno.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import ar.com.trivoli.gestionturnos.calendario.model.Calendario;
import ar.com.trivoli.gestionturnos.turno.model.Turno;
import ar.com.trivoli.gestionturnos.turno.repository.ITurnoRepository;

/**
 * @author ramirez
 *
 */

@Service
@Transactional
public class TurnoService {

	@Autowired
	private ITurnoRepository turnoRepository;
	
	public List<Turno> buscarTurnosPorCalendario (Calendario calendario) {
		//TODO ver si es viable un metodo que busque por id de calendario en lugar de objeto calendario
		
		return turnoRepository.findByCalendario(calendario);
	}
}
