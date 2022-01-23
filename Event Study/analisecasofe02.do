cls
clear

capture log close

set linesize 90
set scheme s2mono

graph drop _all

cd "C:\Users\Legion\OneDrive - Universidade de Aveiro\Universidade\4_ano\1_semestre\FA\trabalho2\"

local mysheets BCP //BES BPI BRISA CIMPOR EDP JM PT SEMAPA SONAE ZON SONAECOM EDPRENOV GALP MOTAENGIL PORTUCEL REN SONAEIND ALTRI COFINA IMPRESA MEDIACAP NOVABASE PARAREDE TEIXEIRADUARTE

foreach sheetname of local mysheets {
	import excel "excel_tratado.xlsx", sheet(`sheetname') firstrow allstring
	display "`sheetname'"

	destring PSI20, generate(psi20_precos)

	destring INDICE_PRECO, generate(indice_preco)
	destring EVENTOS_GERAL, generate(eventos_geral)

	drop PSI20
	drop PSI20VOL
	drop PSI20CAP
	drop PSI20RETORNOS
	drop EVENTOS_RESULTADOS
	drop EVENTOS_DIVIDENDOS
	drop DATA_RESULTADOS
	drop DATA_DIVIDENDOS
	drop INDICE_PRECO
	drop INDICE_VOL
	drop INDICE_CAP
	drop RETORNOS
	drop EVENTOS_GERAL

	gen ln_psi20 = ln(psi20_precos)
	gen r_psi20 = ln_psi20 - ln_psi20[_n-1]
	drop ln_psi20

	gen ln_indice = ln(indice_preco)
	gen r_indice = ln_indice - ln_indice[_n-1]
	drop ln_indice

	//drop if r_indice == .

	//gen data_resultados = date(DATA, "DMY")
	//format %td data_resultados 

	//drop if eventos_geral == 1

	//gen time = _n
	//tsset time

	//twoway (line retornos time), ytitle(Retornos) xtitle(Tempo) title(Retornos ao longo do tempo)

	//histogram retornos, bin(70) normal ytitle(Densidade) xtitle(Retornos) xscale(range(-0.1 0.1))

	su r_indice

	//regress r_indice r_psi20
	
	//clear
}
/*

import excel "excel_tratado.xlsx", sheet("BCP") firstrow allstring


destring PSI20, generate(psi20_precos)

destring INDICE_PRECO, generate(indice_preco)
destring EVENTOS_GERAL, generate(eventos_geral)

drop PSI20
drop PSI20VOL
drop PSI20CAP
drop PSI20RETORNOS
drop EVENTOS_RESULTADOS
drop EVENTOS_DIVIDENDOS
drop DATA_RESULTADOS
drop DATA_DIVIDENDOS
drop INDICE_PRECO
drop INDICE_VOL
drop INDICE_CAP
drop RETORNOS
drop EVENTOS_GERAL

gen ln_psi20 = ln(psi20_precos)
gen r_psi20 = ln_psi20 - ln_psi20[_n-1]

drop ln_psi20

gen ln_indice = ln(indice_preco)
gen r_indice = ln_indice - ln_indice[_n-1]

drop ln_indice

drop if r_indice == .

gen data_resultados = date(DATA, "DMY")
format %td data_resultados 

drop if eventos_geral == 1

gen time = _n
tsset time

//twoway (line retornos time), ytitle(Retornos) xtitle(Tempo) title(Retornos ao longo do tempo)

//histogram retornos, bin(70) normal ytitle(Densidade) xtitle(Retornos) xscale(range(-0.1 0.1))

//su r_indice

regress r_indice r_psi20