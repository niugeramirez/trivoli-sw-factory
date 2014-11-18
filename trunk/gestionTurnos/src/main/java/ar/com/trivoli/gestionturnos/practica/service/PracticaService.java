/**
 * 
 */
package ar.com.trivoli.gestionturnos.practica.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import ar.com.trivoli.gestionturnos.practica.model.Practica;
import ar.com.trivoli.gestionturnos.practica.repository.IPracticaRepository;

/**
 * @author ramirez
 *
 */

@Service
@Transactional
public class PracticaService {
	
	@Autowired
	private IPracticaRepository practicaRepository;
	/************************************************************************************************************************************************************************/
	private Sort ordenPredeterminado() {
		return new Sort(Sort.Direction.ASC, "titulo");
	}
	
	/************************************************************************************************************************************************************************/
	@Transactional(readOnly = true)
	public List<Practica> recuperarTodos() {

		return (List<Practica>) practicaRepository.findAll(ordenPredeterminado());
	}
	/************************************************************************************************************************************************************************/
	
}
