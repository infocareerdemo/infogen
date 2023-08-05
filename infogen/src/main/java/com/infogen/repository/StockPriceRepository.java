package com.infogen.repository;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.CrudRepository;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import com.infogen.entity.StockPrice;

@Repository
public interface StockPriceRepository extends JpaRepository<StockPrice, Integer>, CrudRepository<StockPrice, Integer> {

	@Query(value = "select * from stock_price sp where spsymbol =:spsymbol and spinstrument =:spinstrument", nativeQuery = true)
	StockPrice findBySpsymbolAndSpinstrument(@Param(value = "spsymbol") String spsymbol,
			@Param(value = "spinstrument") String spinstrument);

}
