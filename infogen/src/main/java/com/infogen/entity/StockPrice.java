package com.infogen.entity;

import java.util.Date;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;

@Data
@Entity
@Table(name = "stock_price")
public class StockPrice {
	
	@Id
	private int spid;
	@Column(name="spisin")
	private String spisin;
	@Column(name="spinstrument")
	private String spinstrument;
	@Column(name="spsymbol")
	private String spsymbol;
	@Column(name="spexpiry_dt")
	private String spexpiry_dt;
	
	@Column(name="tdopen")
	private Float tdopen;
	@Column(name="tdhigh")
	private Float tdhigh;
	@Column(name="tdlow")
	private Float tdlow;
	@Column(name="tdclose")
	private Float tdclose;
	@Column(name="tdtimestamp")
	private String tdtimestamp;
	
	@Column(name="twopen")
	private Integer twopen;
	@Column(name="twhigh")
	private Integer twhigh;
	@Column(name="twlow")
	private Integer twlow;
	@Column(name="twclose")
	private Integer twclose;
	@Column(name="twtimestamp")
	private String twtimestamp;
	
	@Column(name="tmopen")
	private Integer tmopen;
	@Column(name="tmhigh")
	private Integer tmhigh;
	@Column(name="tmlow")
	private Integer tmlow;
	@Column(name="tmclose")
	private Integer tmclose;
	@Column(name="tmtimestamp")
	private String tmtimestamp;
	
	@Column(name="tyopen")
	private Integer tyopen;
	@Column(name="tyhigh")
	private Integer tyhigh;
	@Column(name="tylow")
	private Integer tylow;
	@Column(name="tyclose")
	private Integer tyclose;
	@Column(name="tytimestamp")
	private String tytimestamp;
	
	@Column(name="pdopen")
	private Integer pdopen;
	@Column(name="pdhigh")
	private Integer pdhigh;
	@Column(name="pdlow")
	private Integer pdlow;
	@Column(name="pdclose")
	private Integer pdclose;
	@Column(name="pdtimestamp")
	private String pdtimestamp;
	
	
	@Column(name="pwopen")
	private Integer pwopen;
	@Column(name="pwhigh")
	private Integer pwhigh;
	@Column(name="pwlow")
	private Integer pwlow;
	@Column(name="pwclose")
	private Integer pwclose;
	@Column(name="pwtimestamp")
	private String pwtimestamp;
	
	@Column(name="pmopen")
	private Integer pmopen;
	@Column(name="pmhigh")
	private Integer pmhigh;
	@Column(name="pmlow")
	private Integer pmlow;
	@Column(name="pmclose")
	private Integer pmclose;
	@Column(name="pmtimestamp")
	private String pmtimestamp;
	
	@Column(name="pyopen")
	private Integer pyopen;
	@Column(name="pyhigh")
	private Integer pyhigh;
	@Column(name="pylow")
	private Integer pylow;
	@Column(name="pyclose")
	private Integer pyclose;
	@Column(name="pytimestamp")
	private String pytimestamp;
	
	
	@Column(name="slastupdatedate")
	private Date slastupdatedate;
	
	
}
	

