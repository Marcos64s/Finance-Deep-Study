cls
clear

capture log close

set linesize 90
set scheme s2mono

graph drop _all

// altera o path para o da pasta onde se encontra o script
local path = ""

if "`path'" == "" {
	di "Alterar a variavel path para onde esta o excel e onde quer guardar os resultados"
	exit
}

cd "`path'"

// Importar os dados

import excel "dowjones.xlsx", firstrow allstring

// se save = 1 grava todos os resultados

local save = 1

// Transformar os dados importados em floats

destring DJ, generate(precos)
drop DJ
destring M3TBILL, generate(m3tbill)
drop M3TBILL

// calculo dos retornos

gen ln_indice = ln(precos)
gen retornos = (ln_indice - ln_indice[_n-1]) * 100
drop ln_indice
drop if retornos == .

// Criar a variavel temporal

gen data = date(DATA, "MDY")
format %td data
gen time = _n
tsset time

// Criar as series 2-day period, 3-day period, 4-day period e semanal

gen dias_semana = dow(data)
gen helper = _n

foreach i of numlist 2 3 4 {
	gen periodo_`i' = .
	replace periodo_`i' = 1 if mod(helper, `i') == 0
} 

gen semanal = 1 if dias_semana == 1

drop helper
drop dias_semana

// gráficos e estatisticas -> Figura 1, 2 e 3

// Figura_1
line precos data, ylabel(0(2000)12000) yline(2000, lpattern(--)) ytitle(Preço) xtitle(Tempo) title(Preços ao longo do tempo) name(Figura_1)

// Figura_2
line retornos data, ylabel(-8(2)6) ytitle(Retornos) xtitle(Tempo) title(Retornos ao longo do tempo) name(Figura_2)

// histograma_retornos
histogram retornos, normal ytitle(Densidade) xtitle(Retornos) title(Histograma de Retornos) name(histograma_retornos)

// Figura_3_A
ac retornos, ylabel(-0.06(0.02)0.06) xlabel(0(2)20) lags(20) ytitle(Correlação) title(Correlograma de retornos) name(Figura_3_A)

// Figura_3_B
gen retornos_2 = retornos ^ 2
ac retornos_2, ylabel(0(0.05)0.25) xlabel(0(2)20) lags(20) ytitle(Correlação) title(Correlograma de retornos quadrados) name(Figura_3_B)

// JB test -> Tabela 1

tabstat retornos, stat(n mean sd min max skewness kurtosis) save
matrix ed = r(StatTotal)'
sktest retornos
local jb = (ed[1,1] / 6) * (ed[1,6] ^ 2 + 0.25 * (ed[1,7] - 3)^2)
local jb_pvalue = r(p_chi2)
display as text "Jarque-Bera: " as result `jb'
matrix a = (`jb',`jb_pvalue')
matrix ed = ed, a
matrix colnames ed = N Média Desvio-Padrão Mínimo Máximo Assimetria Achatamento Jarque-Bera p-value

// Seleção dos lags para os modelos

varsoc retornos time

// GARCH(1,1) -> Figura 4 e Tabela 2, 6

local i = 0

foreach freq in "diarios" "2_diarios" "3_diarios" "4_diarios" "semanal" {
	
	preserve
	
	if "`freq'" == `"2_diarios"' {
		keep if periodo_2 == 1
		}
	else if "`freq'" == `"3_diarios"' {
		keep if periodo_3 == 1
		}
	else if "`freq'" == `"4_diarios"' {
		keep if periodo_4 == 1
		}
	else if "`freq'" == `"semanal"' {
		keep if semanal == 1
		}
	
	gen ln_indice = ln(precos)
	replace retornos = (ln_indice - ln_indice[_n - 1]) * 100
	drop ln_indice
	drop if retornos == .
	
	replace time = _n
	tsset time
	
	quietly: arch retornos, arch(1) garch(1) save
	matrix garch_11 = r(table)
	matlist garch_11
	local b = garch_11[1, 1]
	local alfa = garch_11[1, 2]
	local beta = garch_11[1, 3]
	local omega = garch_11[1, 4]
	local b_se = garch_11[2, 1]
	local alfa_se = garch_11[2, 2]
	local beta_se = garch_11[2, 3]
	local omega_se = garch_11[2, 4]
	
	if (`i' == 0) {
		// Tabela 6
		matrix helper_matrix = (`b', `b_se' \ `omega', `omega_se' \ `alfa', `alfa_se' \ `beta', `beta_se')
		matrix garch_11_results = (`b' \ `omega' \ `alfa' \ `beta')
	}
	else {
		// Tabela 2
		matrix garch_11_results = garch_11_results, (`b' \ `omega' \ `alfa' \ `beta')
	}
	
	if "`freq'" == `"diarios"' {
		predict myVariances_garch, variance
		gen desvio_garch = sqrt(myVariances_garch) * sqrt(252)
		// Figura_4
		line desvio_garch data, ylabel(5(5)35) ytitle(Volatilidade anual) ttitle(Tempo) name(Figura_4)
		matrix colnames helper_matrix = Coeficientes Erro-padrão
		matrix rownames helper_matrix = Constante omega alfa beta
	}
	
	local i = `i' + 1
	
	restore
}

matrix colnames garch_11_results = "Dados diários" "Período de 2 dias" "Período de 3 dias" "Período de 4 dias" "Dados semanais"
matrix rownames garch_11_results = Constante omega alfa beta

// TARCH(1,1,1) - Tabela 4

quietly: arch retornos, arch(1) garch(1) tarch(1) save
matrix tarch_111 = r(table)
matlist tarch_111
local b = tarch_111[1, 1]
local gama = tarch_111[1, 3]
local alfa = tarch_111[1, 2]
local beta = tarch_111[1, 4]
local omega = tarch_111[1, 5]
local b_se = tarch_111[2, 1]
local gama_se = tarch_111[2, 3]
local alfa_se = tarch_111[2, 2]
local beta_se = tarch_111[2, 4]
local omega_se = tarch_111[2, 5]
matrix tarch_111_results = (`b', `b_se' \ `omega', `omega_se' \ `alfa', `alfa_se' \ `gama', `gama_se' \ `beta', `beta_se')
matrix colnames tarch_111_results = Coeficientes Erro-padrão
matrix rownames tarch_111_results = Constante omega alfa gama beta
predict myVariances_tarch, variance
gen desvio_tarch_1 = sqrt(myVariances_tarch) * sqrt(252)

line desvio_tarch_1 data, ylabel(5(5)35) ytitle(Volatilidade anual) ttitle(Tempo) name(volatilidade_tarch)

// GARCH(1,1) - X -> Tabela 5

quietly: arch retornos, arch(1) garch(1) het(L.m3tbill)
matrix garch_11_X = r(table)
matlist garch_11_X
local b = garch_11_X[1, 1]
local phi = garch_11_X[1, 2]
local omega = garch_11_X[1, 3]
local alfa = garch_11_X[1, 4]
local beta = garch_11_X[1, 5]
local b_se = garch_11_X[2, 1]
local phi_se = garch_11_X[2, 2]
local omega_se = garch_11_X[2, 3]
local alfa_se = garch_11_X[2, 4]
local beta_se = garch_11_X[2, 5]
matrix garch_11_X_results = (`b', `b_se' \ `omega', `omega_se' \ `alfa', `alfa_se' \ `beta', `beta_se' \ `phi', `phi_se')
matrix colnames garch_11_X_results = Coeficientes Erro-padrão
matrix rownames garch_11_X_results = Constante omega alfa beta phi
predict myVariances_garch_X, variance
gen desvio_garch_X = sqrt(myVariances_garch_X) * sqrt(252)

line desvio_garch_X data, ylabel(5(5)35) ytitle(Volatilidade anual) ttitle(Tempo) name(volatilidade_garch_X)

// verificar se as pastas para guardar os graficos e os excels de resultados existem
// guardar tudo

if `save' == 1 {
	
	capture mkdir graficos
	capture mkdir excels
	
	graph export "graficos/Figura 1.png", replace name(Figura_1)
	graph export "graficos/Figura 2.png", replace name(Figura_2)
	graph export "graficos/histograma_retornos.png", replace name(histograma_retornos)
	graph export "graficos/Figura 3 A.png", replace name(Figura_3_A)
	graph export "graficos/Figura 3 B.png", replace name(Figura_3_B)
	graph export "graficos/volatilidade_garch_X.png", replace name(volatilidade_garch_X)
	graph export "graficos/volatilidade_tarch.png", replace name(volatilidade_tarch)
	graph export "graficos/Figura 4.png", replace name(Figura_4)

	putexcel set "excels/Tabela 1.xlsx", replace
	putexcel A1 = matrix(ed), names
	putexcel set "excels/Tabela 2.xlsx", replace
	putexcel A1 = matrix(helper_matrix), names
	putexcel set "excels/Tabela 4.xlsx", replace
	putexcel A1 = matrix(tarch_111_results), names
	putexcel set "excels/Tabela 5.xlsx", replace
	putexcel A1 = matrix(garch_11_X_results), names
	putexcel set "excels/Tabela 6", replace
	putexcel A1 = matrix(garch_11_results), names
	
}

// Print Tabelas

matlist ed
matlist helper_matrix
matlist tarch_111_results
matlist garch_11_results
matlist garch_11_X_results
matlist garch_11_results
