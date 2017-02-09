* Change log:
* version 0.0.2 02mar2016
* major rewrite in this version. Since the aim is to write tab2xl2 so it generates excel files almost ready for publishing, I decided to changed the table presentation. Now column or row percentages are shown side by side of frequencies, as opposed to under it. Also corrected a bug pointed by Bryan.
*
* version 0.0.1 12jan2016
* based on tab2xl http://blog.stata.com/2013/09/25/export-tables-to-excel/


program tab2xl2
	version 13.1
	syntax varlist(min=2 max=2) using/, row(integer) col(integer) [replace modify colf rowf withouthead missing]


	* Use the arguments to define the variables to be used in tabulation 
	local varRow `1' /* points to the variable to be displayed in the row */
	local varCol `2' /* points to the variable to be dipslayed in the column */
	local rowStart = `row' /* row number where to start plotting the data */
	local colStart = `col' /* column number where to start plotting the data */

	quietly tabulate `varRow' `varCol', matcell(freq) matrow(varRowNames) matcol(varColNames) `missing' /* this is the command that tab2xl2 is using "under the hood" to tabulate */

	local observationsNum = r(N) /* saves total number of observations */

	* saves rows and columns VARIABLES LABELS
	local varRowLabel : variable label `varRow'
	if "`varRowLabel'" == "" local varRowLabel = "`varRow'"
	local varColLabel : variable label `varCol'
	if "`varColLabel'" == "" local varColLabel = "`varCol'"


	* test if user chose more than one of the following options.
	if "`colf'"=="colf" & "`rowf'"=="rowf" {
		display as error "options rowf and colf "
		error 184
	}



	* prepares to use putexcel for plotting the data
	quietly putexcel set `"`using'"', `replace' `modify' keepcellformat


	* Start plotting the table. First row with variable labels or names
	local column1 = `colStart'
	num2base26 `column1'
	local columnLetter1 "`r(col_letter)'"
	local column2 = `colStart' + 1
	num2base26 `column2'
	local columnLetter2 "`r(col_letter)'"


	if "`withouthead'"=="withouthead" quietly putexcel `columnLetter1'`rowStart'=("`varRowLabel'")

	else quietly putexcel `columnLetter1'`rowStart'=("`varRowLabel'") `columnLetter2'`rowStart'=("`varColLabel'")


	* Second row: display columns VALUES LABELS. If rowf or colf options selected, third row will display "N" and "%" headers.
	if "`withouthead'"=="" {

		local columns = colsof(varColNames)
		local column = `colStart' + 1
		local row = `rowStart' + 1

		if "`rowf'"=="rowf" | "`colf'"=="colf" local steps = 2 /* if user chooses one of the mentioned options, jump two columns to the right instead of one when plotting variable labels. This is because next row will have "N"s and "%"s. */
		else local steps = 1 

		forvalues i = 1 / `columns' {

			local val = varColNames[1,`i']
			local val_lab : label (`varCol') `val'
			num2base26 `column'
			local columnLetter "`r(col_letter)'"
			quietly putexcel `columnLetter'`row'=("`val_lab'")
			sleep 10

			if "`rowf'"=="rowf" | "`colf'"=="colf" { /* plot the "N"s and "%"s */
				local row = `row' + 1
				local percentCol = `column'+1
				num2base26 `percentCol'
				local percentColLetter "`r(col_letter)'"

				quietly putexcel `columnLetter'`row'=("N") `percentColLetter'`row'=("%")
				local row = `row' - 1
			}

			local column = `column' + `steps'
		}

		num2base26 `column'
		local columnLetter "`r(col_letter)'"
		quietly putexcel `columnLetter'`row'=("Total") /* Last column contains text "Total" */
		sleep 10
		if "`rowf'"=="rowf" | "`colf'"=="colf" { /* plot the "N"s and "%"s */
			local row = `row' + 1
			local percentCol = `column'+1
			num2base26 `percentCol'
			local percentColLetter "`r(col_letter)'"

			quietly putexcel `columnLetter'`row'=("N") `percentColLetter'`row'=("%")
		}
	}


	* Third row: display rows VALUES LABELS. If rowf or colf options selected, fourth row instead. If withouthead, subtract one or two rows.
	local column = `colStart'
	num2base26 `column'
	local columnLetter "`r(col_letter)'"	
	local rows = rowsof(varRowNames)

	if "`rowf'"=="rowf" | "`colf'"=="colf" local row = `rowStart' + 3
	else local row = `rowStart' + 2

	if "`withouthead'"=="withouthead" local row = `row' - 1
	if "`withouthead'"=="withouthead" & ("`rowf'"=="rowf" | "`colf'"=="colf") local row = `row' - 1

	forvalues i = 1/`rows' {
		local val = varRowNames[`i',1]
		local val_lab : label (`varRow') `val'

		quietly putexcel `columnLetter'`row'=("`val_lab'")
		sleep 10
		
		local row = `row' + 1
	}



	* Calculate column and row totals
	matrix U = J(1,rowsof(freq),1)
	matrix columnsTotals = U*freq

	matrix V = J(colsof(freq),1,1)
	matrix rowsTotals = freq*V




	* Fill in frequencies and rows totals. If colf or rowf options, also fill in percentages.
	if "`rowf'"=="rowf" | "`colf'"=="colf" local row = `rowStart' + 3
	else local row = `rowStart' + 2

	if "`withouthead'"=="withouthead" local row = `row' - 1
	if "`withouthead'"=="withouthead" & ("`rowf'"=="rowf" | "`colf'"=="colf") local row = `row' - 1	

	local rows = rowsof(freq)
	local cols = colsof(freq)


	forvalues i = 1/`rows' {

		local column = `colStart' + 1

		forvalues j = 1/`cols' { /* frequencies and percentages */

			num2base26 `column'
			local columnLetter "`r(col_letter)'"
			local percentCol = `column' + 1
			num2base26 `percentCol'
			local percentColLetter "`r(col_letter)'"

			local freq_val = freq[`i',`j']

			if "`colf'" == "colf" local percent_val = `freq_val'/columnsTotals[1,`j']*100
			else if "`rowf'" == "rowf" local percent_val = `freq_val'/rowsTotals[`i',1]*100
			local percent_val : display %9.1fc `percent_val'

			if "`colf'" == "colf" | "`rowf'" == "rowf" {
				quietly putexcel `columnLetter'`row'=(`freq_val') `percentColLetter'`row'=(`percent_val')
				sleep 10
				local column = `column' + 2
			}
			else {
				quietly putexcel `columnLetter'`row'=(`freq_val')
				sleep 10
				local column = `column' + 1	
			}
		}


		/* row totals (right hand side) */
		num2base26 `column'
		local columnLetter "`r(col_letter)'"
		local percentCol = `column' + 1
		num2base26 `percentCol'
		local percentColLetter "`r(col_letter)'"
		if "`colf'" == "colf" local percentTotal = rowsTotals[`i',1]/`observationsNum'*100 /* todo: algo errado */
		else if "`rowf'" == "rowf" local percentTotal = 100
		local percentTotal : display %9.1fc `percentTotal'


		if "`colf'" == "colf" | "`rowf'" == "rowf" {
			quietly putexcel `columnLetter'`row'=(rowsTotals[`i',1]) `percentColLetter'`row'=(`percentTotal')
			sleep 10
		}
		else {
			quietly putexcel `columnLetter'`row'=(rowsTotals[`i',1])
			sleep 10
		}


		local row = `row' + 1		
	}




	* Fill in column totals (bottom)
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
		local percentCol = `column' + 1
		num2base26 `percentCol'
		local percentColLetter "`r(col_letter)'"
		if "`rowf'"=="rowf" local percentTotal = columnsTotals[1,`i']/`observationsNum'*100 /*todo: algo de errado aqui*/
		else if "`colf'"=="colf" local percentTotal = 100
		local percentTotal : display %9.1fc `percentTotal'

		if "`rowf'"=="rowf" | "`colf'"=="colf" {
			quietly putexcel `columnLetter'`row'=(columnsTotals[1,`i']) `percentColLetter'`row'=(`percentTotal')
			sleep 10
			local column = `column' + 1
		}
		else {
			quietly putexcel `columnLetter'`row'=(columnsTotals[1,`i'])
			sleep 10	
		}


		local column = `column' + 1
	}

	num2base26 `column'
	local columnLetter "`r(col_letter)'"
	local percentCol = `column' + 1
	num2base26 `percentCol'
	local percentColLetter "`r(col_letter)'"

	if "`rowf'"=="rowf" | "`colf'"=="colf" {
		quietly putexcel `columnLetter'`row'=(`observationsNum') `percentColLetter'`row'=(100)
	}
	else {
		quietly putexcel `columnLetter'`row'=(`observationsNum')
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
