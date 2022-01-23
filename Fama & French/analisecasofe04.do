cls
clear

capture log close

set linesize 90
set scheme s2mono

graph drop _all

// altera o path para o da pasta onde se encontra o script
local path = "D:/OneDrive/OneDrive - Universidade de Aveiro/Universidade/4ºAno1ºSemestre/FA/trabalho4"

if "`path'" == "" {
	di "Alterar a variavel path para onde esta o excel e onde quer guardar os resultados"
	exit
}

cd "`path'"

// Importar os dados

import excel "DADOS1963.xlsx", firstrow allstring

// se save = 1 grava todos os resultados

local save = 1

// Transformar os dados importados em floats
//
// destring me1bm1, generate(precos)
// drop me1bm1
// destring me2bm1, generate(precos2)
// drop me2bm1

// calculo dos retornos

// gen ln_indice = ln(precos)
// gen retornos = (ln_indice - ln_indice[_n-1]) * 100
// drop ln_indice
// drop if retornos == .

// Criar a variavel temporal

gen data = date(DATA, "YM")
format %td data
gen time = _n
tsset time


///TABELA 2

// Média
ds
return list
local names = r(varlist)
foreach var of varlist DATA-rf{
	destring `var', replace 
}


//Corelação de cada carteira com o mercado
/*foreach var of varlist me1bm1-me5bm5{
	correlate `var' rmrf, covariance
}*/

set graphics off

gen rm_1 = rmrf + rf
foreach var of varlist rmrf-rm_1{
		
	//lags
	quietly: ac `var', lags(12) generate(`var'lag)
	local lag1 = `var'lag[1]
	local lag2 = `var'lag[2]
	local lag12 = `var'lag[12]
	// test t
	ttest `var' = 0
	return list
	local tt = r(t)
	local medi = r(mu_1)
	local st = r(sd_1)
	if `var' == rmrf {
		matrix tabela_2_1 = (`medi',`st',`tt',`lag1',`lag2',`lag12')
		}
	else {
		matrix tabela_2_1 = tabela_2_1 \ (`medi',`st',`tt',`lag1',`lag2',`lag12')
		}
}
matlist tabela_2_1
putexcel set "Excel/t2_1.xlsx", replace 
	putexcel A1 = matrix(tabela_2_1), names

corr rmrf-hml
return list
matlist r(C)
putexcel set "Excel/correl.xlsx", replace 
	putexcel A1 = matrix(r(C)), names 


foreach x of varlist me1bm1-me5bm5{
	gen rf`x' = `x' - rf
}

foreach x of varlist rfme1bm1-rfme5bm5{
	sum `x'
	return list
	local medi = r(mean)
	local st = r(sd)
	ttest `x' = 0
	local tt = r(t)
	di `medi' `st' `tt'
	if `x' == rfme1bm1 {
		matrix tabela_2 = (`medi',`st',`tt')
		}
	else {
		matrix tabela_2 = tabela_2 \ (`medi',`st',`tt')
		}
}
matlist tabela_2
putexcel set "Excel/t2_2.xlsx", replace 
	putexcel A1 = matrix(tabela_2), names

///TABELA 4
	
foreach x of varlist rfme1bm1-rfme5bm5{
	reg `x' rmrf
	return list
	matrix tables = r(table)
	matrix list tables
	local h1 = tables[1,1]
	local tb = tables[3,1]
	local regs = e(r2_a)
	local ses = e(rmse)
	if `x' == rfme1bm1 {
		matrix tabela_4 = (`h1',`tb',`ses',`regs')
		}
	else {
		matrix tabela_4 = tabela_4 \ (`h1',`tb',`ses',`regs')
		}
}
matlist tabela_4
putexcel set "Excel/t4_1.xlsx", replace 
	putexcel A1 = matrix(tabela_4), names

foreach x of varlist rfme1bm1-rfme5bm5{
	reg `x' rmrf smb hml
	return list
	matrix tables = r(table)
	matrix list tables
	local b1 = tables[1,1]
	local h1 = tables[1,3]
	local s1 = tables[1,2]
	local tb = tables[3,1]
	local th = tables[3,3]
	local ts = tables[3,2]
	local regs = e(r2_a)
	local ses = e(rmse)
	if `x' == rfme1bm1 {
		matrix tabela_5 = (`b1',`h1',`s1',`tb',`th',`ts',`ses',`regs')
		}
	else {
		matrix tabela_5 = tabela_5 \ (`b1',`h1',`s1',`tb',`th',`ts',`ses',`regs')
		}
}
matlist tabela_5
putexcel set "Excel/t5_1.xlsx", replace 
	putexcel A1 = matrix(tabela_5), names


cls
clear

capture log close

set linesize 90
set scheme s2mono

graph drop _all
clear
import excel "DADOS1964.xlsx", firstrow allstring

// se save = 1 grava todos os resultados

local save = 1

// Transformar os dados importados em floats
//
// destring me1bm1, generate(precos)
// drop me1bm1
// destring me2bm1, generate(precos2)
// drop me2bm1

// calculo dos retornos

// gen ln_indice = ln(precos)
// gen retornos = (ln_indice - ln_indice[_n-1]) * 100
// drop ln_indice
// drop if retornos == .

// Criar a variavel temporal

gen data = date(DATA, "YM")
format %td data
gen time = _n
tsset time


///TABELA 2

// Média
ds
return list
local names = r(varlist)
foreach var of varlist DATA-rf{
	destring `var', replace 
}


//Corelação de cada carteira com o mercado
/*foreach var of varlist me1bm1-me5bm5{
	correlate `var' rmrf, covariance
}*/

set graphics off

gen rm_1 = rmrf + rf
foreach var of varlist rmrf-rm_1{
		
	//lags
	quietly: ac `var', lags(12) generate(`var'lag)
	local lag1 = `var'lag[1]
	local lag2 = `var'lag[2]
	local lag12 = `var'lag[12]
	// test t
	ttest `var' = 0
	return list
	local tt = r(t)
	local medi = r(mu_1)
	local st = r(sd_1)
	if `var' == rmrf {
		matrix tabela_2_1 = (`medi',`st',`tt',`lag1',`lag2',`lag12')
		}
	else {
		matrix tabela_2_1 = tabela_2_1 \ (`medi',`st',`tt',`lag1',`lag2',`lag12')
		}
}
matlist tabela_2_1
putexcel set "Excel/t2_1 new.xlsx", replace 
	putexcel A1 = matrix(tabela_2_1), names

corr rmrf-hml
return list
matlist r(C)
putexcel set "Excel/correl new.xlsx", replace 
	putexcel A1 = matrix(r(C)), names 


foreach x of varlist me1bm1-me5bm5{
	gen rf`x' = `x' - rf
}

foreach x of varlist rfme1bm1-rfme5bm5{
	sum `x'
	return list
	local medi = r(mean)
	local st = r(sd)
	ttest `x' = 0
	local tt = r(t)
	di `medi' `st' `tt'
	if `x' == rfme1bm1 {
		matrix tabela_2 = (`medi',`st',`tt')
		}
	else {
		matrix tabela_2 = tabela_2 \ (`medi',`st',`tt')
		}
}
matlist tabela_2
putexcel set "Excel/t2_2 new.xlsx", replace 
	putexcel A1 = matrix(tabela_2), names

///TABELA 4
	
foreach x of varlist rfme1bm1-rfme5bm5{
	reg `x' rmrf
	return list
	matrix tables = r(table)
	matrix list tables
	local h1 = tables[1,1]
	local tb = tables[3,1]
	local regs = e(r2_a)
	local ses = e(rmse)
	if `x' == rfme1bm1 {
		matrix tabela_4 = (`h1',`tb',`ses',`regs')
		}
	else {
		matrix tabela_4 = tabela_4 \ (`h1',`tb',`ses',`regs')
		}
}
matlist tabela_4
putexcel set "Excel/t4_1 new.xlsx", replace 
	putexcel A1 = matrix(tabela_4), names

foreach x of varlist rfme1bm1-rfme5bm5{
	reg `x' rmrf smb hml
	return list
	matrix tables = r(table)
	matrix list tables
	local b1 = tables[1,1]
	local h1 = tables[1,3]
	local s1 = tables[1,2]
	local tb = tables[3,1]
	local th = tables[3,3]
	local ts = tables[3,2]
	local regs = e(r2_a)
	local ses = e(rmse)
	if `x' == rfme1bm1 {
		matrix tabela_5 = (`b1',`h1',`s1',`tb',`th',`ts',`ses',`regs')
		}
	else {
		matrix tabela_5 = tabela_5 \ (`b1',`h1',`s1',`tb',`th',`ts',`ses',`regs')
		}
}
matlist tabela_5
putexcel set "Excel/t5_1 new.xlsx", replace 
	putexcel A1 = matrix(tabela_5), names

