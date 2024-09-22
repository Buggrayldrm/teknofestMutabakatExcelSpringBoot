package org.example.excelproject.repository;

import org.example.excelproject.entities.DataEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;


@Repository
public interface ExcelRepository extends JpaRepository<DataEntity, Long> {

}