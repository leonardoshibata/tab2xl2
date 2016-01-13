*! version 0.0.1 12jan2016
* based on tab2xl http://blog.stata.com/2013/09/25/export-tables-to-excel/
program tab2xl2
	version 13.1
	syntax varlist(min=2 max=2) using/, row(integer) col(integer) [replace modify colf rowf]

	* Use the arguments to define the variables to be used in tabulation 
	local varRow `1'
	local varCol `2'
	local rowStart = `row'
	local colStart = `col'

	quietly tabulate `varRow' `varCol', matcell(freq) matrow(varRowNames) matcol(varColNames)

	local observationsNum = r(N)

	local varRowLabel : variable label `varRow'
	if "`varRowLabel'" == "" local varRowLabel = "`varRow'"

	local varColLabel : variable label `varCol'
	if "`varColLabel'" == "" local varColLabel = "`varCol'"

	quietly putexcel set `"`using'"', `replace' `modify' keepcellformat

	* First row with variable labels or names
	local column1 = `colStart'
	num2base26 `column1'
	local columnLetter1 "`r(col_letter)'"
	local column2 = `colStart' + 1
	num2base26 `column2'
	local columnLetter2 "`r(col_letter)'"

	quietly putexcel `columnLetter1'`rowStart'=("`varRowLabel'") `columnLetter2'`rowStart'=("`varColLabel'")

	* Second row with varCol values labels
	local columns = colsof(varColNames)
	local column = `colStart' + 1
	local row = `rowStart' + 1

	forvalues i = 1/`columns' {

		local val = varColNames[1,`i']
		local val_lab : label (`varCol') `val'
		num2base26 `column'
		local columnLetter "`r(col_letter)'"

		quietly putexcel `columnLetter'`row'=("`val_lab'")
		sleep 10
		local column = `column' + 1
	}

	num2base26 `column'
	local columnLetter "`r(col_letter)'"
	quietly putexcel `columnLetter'`row'=("Total")
	sleep 10

	* Starting at third row, varRow values labels
	local rows = rowsof(varRowNames)
	local row = `rowStart' + 2

	forvalues i = 1/`rows' {
		local val = varRowNames[`i',1]
		local val_lab : label (`varRow') `val'

		quietly putexcel A`row'=("`val_lab'")
		sleep 10
		
		if "`colf'" != "" | "`rowf'" != "" local row = `row' + 2
		else local row = `row' + 1
		sleep 10
	}

	* Calculate column and row totals
	matrix U = J(1,rowsof(freq),1)
	matrix columnsTotals = U*freq

	matrix V = J(colsof(freq),1,1)
	matrix rowsTotals = freq*V


	* Fill in the frequencies, row total and percentages
	local row = `rowStart' + 2
	local rows = rowsof(freq)
	local cols = colsof(freq)


	forvalues i = 1/`rows' {

		local column = `colStart' + 1

		forvalues j = 1/`cols' {

			num2base26 `column'
			local columnLetter "`r(col_letter)'"

			local freq_val = freq[`i',`j']

			quietly putexcel `columnLetter'`row'=(`freq_val')
			sleep 10
			local column = `column' + 1
		}

		num2base26 `column'
		local columnLetter "`r(col_letter)'"

		quietly putexcel `columnLetter'`row'=(rowsTotals[`i',1])
		sleep 10


		local row = `row' + 1

		if "`colf'" == "colf" {

			local column = `colStart' + 1

			forvalues j = 1/`cols' {

				num2base26 `column'
				local columnLetter "`r(col_letter)'"

				local freq_val = freq[`i',`j']
				local percent_val = `freq_val'/columnsTotals[1,`j']*100
				local percent_val : display %9.2f `percent_val'

				quietly putexcel `columnLetter'`row'=(`percent_val')
				sleep 10
				local column = `column' + 1
			}
			local percentTotal = rowsTotals[`i',1]/`observationsNum'*100
			local percentTotal : display %9.2f `percentTotal'
			
			num2base26 `column'
			local columnLetter "`r(col_letter)'"

			quietly putexcel `columnLetter'`row'=(`percentTotal')

			local row = `row' + 1
		}

		else if "`rowf'" == "rowf" {

			local column = `colStart' + 1

			forvalues j = 1/`cols' {

				num2base26 `column'
				local columnLetter "`r(col_letter)'"

				local freq_val = freq[`i',`j']
				local percent_val = `freq_val'/rowsTotals[`i',1]*100
				local percent_val : display %9.2f `percent_val'

				quietly putexcel `columnLetter'`row'=(`percent_val')
				sleep 10
				local column = `column' + 1
			}

			num2base26 `column'
			local columnLetter "`r(col_letter)'"

			quietly putexcel `columnLetter'`row'=(100.00)

			local row = `row' + 1
		}
		
	}

	* Fill in column totals
	local columns = colsof(varColNames)
	local column = `colStart'
	num2base26 `column'
	local columnLetter "`r(col_letter)'"

	quietly putexcel `columnLetter'`row'=("Total")
	sleep 10

	local column = `column' + 1

	forvalues i = 1/`columns' {

		num2base26 `column'
		local columnLetter "`r(col_letter)'"

		quietly putexcel `columnLetter'`row'=(columnsTotals[1,`i'])
		sleep 10

		local column = `column' + 1
	}

	num2base26 `column'
	local columnLetter "`r(col_letter)'"

	quietly putexcel `columnLetter'`row'=(`observationsNum')
	sleep 10

	if "`colf'" == "colf" {

		local row = `row' + 1
		local column = `colStart' + 1

		forvalues i = 1/`columns' {

			num2base26 `column'
			local columnLetter "`r(col_letter)'"

			quietly putexcel `columnLetter'`row'=(100.00)
			sleep 10

			local column = `column' + 1
		}

		num2base26 `column'
		local columnLetter "`r(col_letter)'"

		quietly putexcel `columnLetter'`row'=(100.00)
		sleep 10	
	}

	else if "`rowf'" == "rowf" {

		local row = `row' + 1
		local column = `colStart' + 1

		forvalues i = 1/`columns' {

			num2base26 `column'
			local columnLetter "`r(col_letter)'"
			local percentTotal = columnsTotals[1,`i']/`observationsNum'*100
			local percentTotal : display %9.2f `percentTotal'

			quietly putexcel `columnLetter'`row'=(`percentTotal')
			sleep 10

			local column = `column' + 1
		}

		num2base26 `column'
		local columnLetter "`r(col_letter)'"

		quietly putexcel `columnLetter'`row'=(100.00)
		sleep 10
	}
end

program num2base26, rclass
	args num

	mata: my_col = strtoreal(st_local("num"))
	mata: col = numtobase26(my_col)
	mata: st_local("col_let", col)
	return local col_letter = "`col_let'"
end

* end of program
