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
	
	Page<Paciente> findByDniStartingWithAndApellidoLikeAndNombreLike	(	Pageable pageable
																			, String dni
																			, String apellido
																			, String nombre
																		);	

	Page<Paciente> findByDniStartingWithAndApellidoLike	(	Pageable pageable
															, String dni
															, String apellido
														);	

	Page<Paciente> findByDniStartingWithAndNombreLike	(	Pageable pageable
															, String dni
															, String nombre
														);

	Page<Paciente> findByNombreLike	(	Pageable pageable
										, String nombre
		);

	Page<Paciente> findByApellidoLike	(	Pageable pageable
											, String apellido
										);

	Page<Paciente> findByApellido		(	Pageable pageable
											, String apellido
		);
}
