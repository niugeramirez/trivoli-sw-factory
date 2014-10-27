/**
 * 
 */
package ar.com.trivoli.gestionturnos.paciente.repository;



import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.paciente.model.Paciente;


/**
 * @author ramirez
 *
 */
public interface IPacienteRepository extends 
		PagingAndSortingRepository<Paciente, Integer>{

	Page<Paciente> findByDniStartingWith(Pageable pageable, String dni);	

}
