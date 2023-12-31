*******************************************************************************
*   Version v0.4: 2023-05-03
*******************************************************************************
*   Dennis Föste-Eggers	
*   
*   German Centre for Higher Education Research and Science Studies (DZHW)
*   Lange Laube 12, 30159 Hannover         
*   Phone: +49-(0)511 450 670-114	
*   
*   E-Mail (1): dennis.foeste@gmail.com 		
*   E-Mail (2): dennis.foeste@outlook.de
*   E-Mail (3): dennis.foeste@gmx.de
*   
*   E-Mail (4): foeste-eggers@dzhw.eu
*******************************************************************************
*   Program name: auto_code.ado     
*   Program purpose: Wrapper for automated code assignment based on excel files
*                    containing stings and associated codes.			
*******************************************************************************
*   Changes made:
*   Version 0.1: added GPL 
*   Version 0.2: adaptation of basic syntax
*   Version 0.3:
* 
*******************************************************************************
*   License: GPL (>= 3)
*     
*   auto_code.ado for Stata
*   Copyright (C) 2023 Foeste-Eggers, Dennis 
*   
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU General Public License as published by
*   the Free Software Foundation, either version 3 of the License, or
*   (at your option) any later version.
*   
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
*   GNU General Public License for more details.
*   
*   You should have received a copy of the GNU General Public License
*   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*   
*******************************************************************************
*   Citation: This code is © D. Foeste-Eggers, 2023, and it is made 
*				 available under the GPL license enclosed with the software.
*
*!			Over and above the legal restrictions imposed by this license, if !
*!          you use this program for any (academic) publication then you are  !
*! 			obliged to provide proper attribution.                            !
*
*   D. Foeste-Eggers auto_code.ado for Stata, v0.3 (2023). 
*           [weblink].
*
*******************************************************************************
cap program drop auto_code 
program define auto_code  , nclass
    version 15
	
    syntax varlist [if] [in] using [, 			///
                                import_excel_options(passthru) ///
                                noinputfile ///  undocumented -tbd
                                suffix_harm(namelist min=1 max=1) /// undocumented 
                                suffix_source(namelist min=1 max=1) /// undocumented 
                                suffix_code(namelist min=1 max=1) /// undocumented 
                                keep_othervar(varlist) /// undocumented 
                                othmergevar(passthru) /// undocumented -tbd
                                ZOFARmissings /// undocumented 
                                reshape_wide /// undocumented -tbd 
								dropharm ///  undocumented 
								dropsource ///  undocumented 
								codeisstring ///  undocumented  
								flagmulti(namelist min=1 max=1) /// undocumented / suffix
								firstcodeofastring  /// going to be outdated
								multiplecodes /// undocumented / no alternatives yet e. g. random / last
								viewinput ///
								///
								/// str_harm options:
								xlsx_replacements(passthru)	 /// handover to str_harm   
                                sheet_replacements(passthru) /// handover to str_harm
								viewreplacements			 /// 
                                ]
                                
                                * GENerate(namelist)	    /// namelist
                                * REPLACE 	            /// replace vars
                                * SUFfix(namelist max=1)  /// undocumented -tbd
                                * PREfix(namelist max=1)  /// undocumented -tbd
                                * MISSINGS(passthru) 		/// undocumented -tbd
                                * MISETs(passthru)        /// undocumented -tbd
                                * XLSxfile(passthru)		/// undocumented -tbd
                                * sheet(passthru)			/// undocumented -tbd
                                * cellrange(passthru)	 // undocumented -tbd  
                                
     
*******************************************************************************

** to do: **
* 
* use [using]
* [if] [in] qualifier
* missings 
* confirm file
* save option
* duplicates within order...
* andere Spalten: Variablen(listen) / auch für mergen 
* ggf. Infos nicht fuer das mergen, aber mitzunehmen
* Spalten aus Listen spezifierbar machen, die genutzt werden sollen
* synonyme Funktion integrieren
* missings als funktion integrieren
* harm funktion flexibler machen

* use rclass returns for coupling with export_mancode

    di as txt "Automated code assignment via auto_code.ado by D. Foeste-Eggers (2023,V0.3)"
    di ""
        
    tempfile syntax_l temp_o
    
    if `"`temp_o'"' ~= `""' save `temp_o' , replace	 

 
if `"`import_excel_options'"' ~=  `""' {
    local import_excel_options = subinstr(`"`import_excel_options'"',`"import_excel_options("',"",1)
    local import_excel_options = substr(`"`import_excel_options'"',1,`=length(`"`import_excel_options'"')-1')
    local import_excel_options = trim(`"`import_excel_options'"')
    if inrange(strpos(`"`import_excel_options'"',`","'),2,`=length(`"`import_excel_options'"')') {
        noi cap di as error `"please check the "," you specified within the option import_excel_options: it may cause an error"'
    }  
    else if strpos(`"`import_excel_options'"',`","')==0 {
        local import_excel_options = `", `import_excel_options'"'
    }
    
    * di `"`import_excel_options'"'
}
	* not needed -- now implemented as options in str_harm 
	* qui if `"`xlsx_replacements'"' ~= `""' {
	* 	local xlsx_replacements = subinstr(`"`xlsx_replacements'"',`"xlsx_replacements("',"",1)
	* 	local xlsx_replacements = substr(`"`xlsx_replacements'"',1,`=length(`"`xlsx_replacements'"')-1')
	* 	local xlsx_replacements = trim(`"`xlsx_replacements'"')
	* 	if `"`sheet_replacements'"' == `""' local sheet_replacements = `"Sheet1"'
	* 	else {
	* 		local sheet_replacements = subinstr(`"`sheet_replacements'"',`"sheet_replacements("',"",1)
	* 		local sheet_replacements = substr(`"`sheet_replacements'"',1,`=length(`"`sheet_replacements'"')-1')
	* 		local sheet_replacements = trim(`"`sheet_replacements'"')
	* 	}
	* }
	
if `"`multiplecodes'"' ~=  `""' | `"`firstcodeofastring'"' ~=  `""' {
	if `"`codeisstring'"' == `""' noi di as error "at least for now multiplecodes option requires codeisstring option"'
	* else 
}
		


if `"`noinputfile'"' ~=  `""' {
    import excel `using' `import_excel_options'
}
else {
    if strpos(`"`import_excel_options'"',`"fristrow"')==0 {
        local import_excel_options = trim(`"`import_excel_options' firstrow"')
    }
    if strpos(`"`import_excel_options'"',`"allsting"')==0 {
        local import_excel_options = trim(`"`import_excel_options' allstring"')
    }
    import excel `using' `import_excel_options' 
    keep if !missing(file) 
    tempvar line
    gen `line' = _n
    sum `line' , meanonly
    if `r(max)' == 0 di as error `"within the input file there are no excel files specified"'
    else {
        local loops = `r(max)'
        * confirm file // pruefen und bereinigen spaeter ergaenzen
    }
    sort order
    
    
    
    gen syntax1 = `"local file"' + order  + "= \`" + `"""' + ///
    trim(file) + `"""' + "'"
    gen syntax2 = `"local sheet"' + order  + "= \`" + `"""' + ///
    trim(sheet) + `"""' + "'"
    gen syntax3 = `"local collect"' + order  + "= \`" + `"""' + ///
    trim(updatecollectionofstringsin) + `"""' + "'"
    gen syntax4 = `"local project"' + order  + "= \`" + `"""' + ///
    trim(projectid) + `"""' + "'"
    
    gen syntax5 = `"local stringcol"' + order  + "= \`" + `"""' + ///
    trim(colofstring) + `"""' + "'"
    gen syntax6 = `"local codecol"' + order  + "= \`" + `"""' + ///
    trim(colofcode) + `"""' + "'"
    gen syntax7 = `"local othcol"' + order  + "= \`" + `"""' + ///
    trim(othcolfromxlsxformerging) + `"""' + "'"
    gen syntax8 = `"local othmergevar"' + order  + "= \`" + `"""' + ///
    trim(othvarfromdtaformerging) + `"""' + "'"
    gen syntax9 = `"local othvar"' + order  + "= \`" + `"""' + ///
    trim(othvarfromdtaforcoding) + `"""' + "'"

    

    keep syntax*
    
    gen num = _n 

    reshape long syntax , i(num) j(s_num)
    sort s_num num
    drop s_num num

    if `"`syntax_l'"'~=`""' outfile syntax using `syntax_l'.do, noquote replace wide
	
	if `"`viewinput'"'~=`""' view `syntax_l'.do
    
    include `syntax_l'.do
    
    use `temp_o', clear
	
	*tempname vl_list1_tn vl_list2_tn


    * use locals if specified
    if `"`keep_othervar'"' ~= `""' local oth_tbk     = `"`keep_othervar'"' // unveraendert
 
    * use defaults, if not specified otherwise
    if `"`suffix_harm'"' == `""' local suffix_harm = `"_harm"'
    if `"`suffix_code'"' == `""' local suffix_code = `"_code"'
    if `"`suffix_sour'"' == `""' local suffix_sour = `"_source"'
    
	*di `"`varlist'"'
	
    local var_list    = `"`varlist'"' // nicht mit VL in Stata arbeiten - einfacher, aber keine stringvars
	
	* vl create `vl_list1_tn' = (`varlist')
	* vl create `vl_list2_tn' = `vl_list1_tn' + (`keep_othvar') 
	
	* novarabbrev {
	* 	foreach var of varlist `var_list' {
	* 		if strpos(`" `oth_tbk' "'
	* 	}
	* }
	*novarabbrev {
		
		keep `oth_tbk' `var_list'
	
	*}
	*keep $(`vl_list2_tn')
    
    tempname str_harm code 
    tempvar merger duplos
    
    forvalues i = 1(1)`loops' {
        tempfile xlsx_codes`i'
        *if `i' ==1 {
            
            preserve
                import excel `file`i'' , sheet(`sheet`i'') clear allstring
                
				
				if `"`stringcol`i''"' == `""' local stringcol`i' = `"A"'
				else local stringcol`i' = upper(`"`stringcol`i''"')
				
				if `"`codecol`i''"'   == `""' local codecol`i'   = `"B"'
				else local codecol`i' = upper(`"`codecol`i''"')
				
                if `"`othcol`i''"'   ~= `""' {
                    local othcol`i'   = subinstr(`"`oth`i''"',`","'," ",.)
                    local othcol`i'   = `" `oth`i'' "'
                }
                
                keep `stringcol`i'' `codecol`i'' `othcol`i'' // othcol prüfen
                rename (`stringcol`i'' `codecol`i'') (str `code')
				if `i' ==1 & `"`viewreplacements'"'~=`""' /// 
					str_harm str , gen(`str_harm') `xlsx_replacements' `sheet_replacements' viewreplacements
                else str_harm str , gen(`str_harm') `xlsx_replacements' `sheet_replacements'
				qui {
					tempvar length 
					gen `length' = length(`str_harm')
					qui sum `length', meanonly
				}
				if `r(max)'<=2045 recast str`r(max)' `str_harm'
                order `str_harm' `code'
                drop str
        
                duplicates drop
                drop if `str_harm' == `""' & `code' == `""'
        
                if `"`codeisstring'"'==`""' {
					cap noi destring `code' , replace // force
					if _rc {
						tempvar destr_code
						destring `code' , gen(`destr_code') force
						sort `str_harm'
						cap noi di as error `"the following codes could not be destringed, and have been ignored:"'
						noi list `str_harm' `code' if missing(`destr_code') & !missing(trim(`code')), sepby(`str_harm') noheader noobs
						* cap noi di as error `"if there are duplicates regarding the strings, this may result in ..."'
						drop `code'
						clonevar `code' = `destr_code'
						drop `destr_code'
					}
				}	
				
                noi duplicates tag `str_harm' , gen(`duplos')
        
                sort `str_harm'
				qui sum `duplos'
                if `r(max)'>0 {
					cap noi di as error  `"the following duplicates in the input file have been ignored:"'
					noi list `str_harm' `code' if `duplos' > 0, sepby(`str_harm') noheader noobs
				}
    
                cap drop if `duplos' > 0 
                cap drop `duplos'
        
        
                if `"`xlsx_codes`i''"' ~= `""' save `xlsx_codes`i'', replace
            restore
        *}
        *else {
        *    preserve
        *        import excel `file`i'' , sheet(`sheet`i'') clear allstring
        *    restore
    }
    
foreach var of varlist `var_list' {
    
    str_harm `var' , gen(`var'`suffix_harm') `xlsx_replacements' `sheet_replacements'
    
    cap drop `str_harm'
    cap drop `code'
    clonevar `str_harm' = `var'`suffix_harm'
	
    qui {
		tempvar length 
		gen `length' = length(`str_harm')
		qui sum `length', meanonly
	}
	if `r(max)'<=2045 recast str`r(max)' `str_harm'
	*recast str256 `str_harm'
    
    forvalues i = 1(1)`loops' {
        *tempfile xlsx_codes`i'
        if `i' ==1 {
            merge m:1 `str_harm'  using `xlsx_codes`i'', gen(`merger')
            gen `var'`suffix_sour' = (`merger'==3)
            sum `var'`suffix_sour' , meanonly
            if `r(max)'== 0 {
                label define `var'`suffix_sour' 0 "" // 1 "SYNTAX" 
            }
            else label define `var'`suffix_sour' 0 "" `i' `"LIST `i'"' 
        
            drop if `merger'==2 // drop nicht benoetigte Codes aus Liste
            drop `merger' 
            
            
        }
        else {
            merge m:1 `str_harm' using `xlsx_codes`i'', gen(`merger') update
            replace `var'`suffix_sour' = `i' if `merger'==4
            sum `merger' , meanonly
            if `r(max)' == 5 {
                di as error`"There are conflicts regarding the assigned codes: "'
                list `str_harm' if `merger' == 5
            }
            drop if `merger'==2 // drop nicht benoetigte Codes aus Codierliste
            drop `merger'
            sum `var'`suffix_sour' , meanonly
            if `r(max)'== `i' label define `var'`suffix_sour' `i' `"LIST `i'"', modify 
        
            if `i' == `loops' cap drop `str_harm'
            
            
            
        }
        
        
    }
    rename `code' `var'`suffix_code'
}

if `"`firstcodeofastring'"' ~=  `""' {
	tempvar wc_var 
	tempname first
	foreach var of varlist `var_list' {
		replace `var'`suffix_code' = " " + `var'`suffix_code' + " "
		replace `var'`suffix_code' = subinstr(`var'`suffix_code',";" ," ",.)
		replace `var'`suffix_code' = subinstr(`var'`suffix_code',":" ," ",.)
		replace `var'`suffix_code' = subinstr(`var'`suffix_code',"," ," ",.)
		replace `var'`suffix_code' = subinstr(`var'`suffix_code',". "," ",.)
		
		gen `wc_var' = wordcount(`var'`suffix_code')
		gen `first'  = (wordcount(`var'`suffix_code')>1)
		
		
		replace `var'`suffix_code' = word(`var'`suffix_code',1) if `first'==1
		replace `var'`suffix_code' = trim(`var'`suffix_code')
		
		destring `var'`suffix_code' , force replace
		
		if `"`flagmulti'"' ~=  `""' rename `first' 	`var'`flagmulti'
		else drop `first'
		
		drop `wc_var'
	}
	
	local codeisstring=`""'
}

if `"`zofarmissings'"' ~= `""'   { 
    foreach var of varlist `var_list' {
        qui {
				if `"`codeisstring'"'==`""' {
					replace `var'`suffix_code' = .u     if trim(`var') == "-9990"
					replace `var'`suffix_code' = .v     if trim(`var') == "-9991"
					replace `var'`suffix_code' = .w     if trim(`var') == "-9992"
					replace `var'`suffix_code' = .x     if trim(`var') == "-9995"
					replace `var'`suffix_code' = .y     if trim(`var') == "-9999"
				}
				else {
					replace `var'`suffix_code' = ".u"   if trim(`var') == "-9990"
					replace `var'`suffix_code' = ".v"   if trim(`var') == "-9991"
					replace `var'`suffix_code' = ".w"   if trim(`var') == "-9992"
					replace `var'`suffix_code' = ".x"   if trim(`var') == "-9995"
					replace `var'`suffix_code' = ".y"   if trim(`var') == "-9999"
				}
                
                sum `var'`suffix_sour' , meanonly
                local mi_sour = `r(max)' + 1
                
                replace `var'`suffix_sour' = `mi_sour'       if trim(`var') == "-9990"
                replace `var'`suffix_sour' = `mi_sour'       if trim(`var') == "-9991"
                replace `var'`suffix_sour' = `mi_sour'       if trim(`var') == "-9992"
                replace `var'`suffix_sour' = `mi_sour'       if trim(`var') == "-9995"
                replace `var'`suffix_sour' = `mi_sour'       if trim(`var') == "-9999"
                
                sum `var'`suffix_sour' , meanonly
                if `r(max)'== `mi_sour' label define `var'`suffix_sour' `mi_sour' "AUTO. MISSING", modify 
                
                
                label values `var'`suffix_sour' `var'`suffix_sour'
        }
                
            
        tab `var'`suffix_sour'
        tab `var'`suffix_sour' if `var'`suffix_sour' ~= `mi_sour'
        
    * if `"`sour_list'"' == `""' local sour_list = `"`var'`suffix_sour'"'
    * else local sour_list = `"`sour_list', `var'`suffix_sour'"'
    
    
        
    }
}  
}

if `"`dropharm'"' ~= `""' foreach var of varlist `var_list' {
			drop `var'`suffix_harm' 
}
if `"`dropsource'"' ~= `""' foreach var of varlist `var_list' {
			drop `var'`suffix_sour' 
}

   
            
          
             

end 




