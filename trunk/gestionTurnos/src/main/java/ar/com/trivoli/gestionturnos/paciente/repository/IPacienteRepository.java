/**
 * 
 */
package ar.com.trivoli.gestionturnos.paciente.repository;

import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.paciente.model.Paciente;

/**
 * @author ramirez
 *
 */
public interface IPacienteRepository extends 
		PagingAndSortingRepository<Paciente, Integer>{

}
