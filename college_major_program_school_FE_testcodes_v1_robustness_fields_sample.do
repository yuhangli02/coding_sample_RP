*Checks how sensitive reduced form wage regressions are to including college and program FEs
*By degree major

use "/home/Projects/MajorChoice/intdata/MajorChoice/majorchoice_wfac_v14f.dta",

sort munic
tab munic

*merges cluster info
merge m:1 munic using /home/Projects/MajorChoice/intdata/EducCollege/School_clusters_allmunic_v2 

tab _merge
tab munic if _merge==2




*OLD:  local controls "college_mom hs_mom hs_mom_miss college_dad hs_dad hs_dad_miss fam_inc73_mom fam_inc73_mom_squared malder malder_miss health1 health2 health_missing schoolave_faminc b1974 b1975 adveng advmath HS_track3 HS_track4"
*UPDATED:  
local controls "fam_inc73_mom college_mom hs_mom hs_mom_miss college_dad hs_dad hs_dad_miss health1 health2 health_missing adveng advmath HS_track3 HS_track4"

foreach var of varlist logwage logpvdispinc {
        * Set the excel output file and output sheet
        putexcel set college_major_program_school_FE_raw_v1_robustness_fields_clusters, modify sheet("Table_FE_`var'")
		*putexcel set "C:\Users\yuhangli\OneDrive - The University of Chicago\Documents\Summer 2023 SISRM RA\Update Tables\robustness_fields_clusters", modify sheet("Table_FE_`var'")
        * Outputting Results
        forvalues f=1/12 {
        sum DGradColl`f' 
        sum cog interpers grit HS_track3 HS_track4 if DGradColl`f'==1
        
        * Regression Model 1
		*reg logwage cog interpers grit `controls' if DGradColl1 == 1 & logwage>-9999
        reg `var' cog interpers grit `controls' if DGradColl`f'==1 & `var'>-9999
		matrix results = r(table)
        matrix a = results[1,1..3]'
        matrix b = results[4,1..3]' 
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for cognitive
		local cell_reference_cog = char(66) + string(`row_number')
		* Create the Excel cell reference for interpersonal
		local cell_reference_int = char(70) + string(`row_number')
		* Create the Excel cell reference for grit
		local cell_reference_grit = char(74) + string(`row_number')
		* Output the b coefficient and p-value 
		matrix cog_baseline = matrix(results[1,1..2]')
		matrix int_baseline = matrix(results[2,1..2]')
		matrix grit_baseline = matrix(results[3,1..2]')
		putexcel `cell_reference_cog' = matrix(cog_baseline)
		putexcel `cell_reference_int' = matrix(int_baseline)
		putexcel `cell_reference_grit' = matrix(grit_baseline) 
        local ra = char(66) + string(45 + `f' - 1)
        putexcel F42 = ("Baseline")  `ra' = (e(r2_a))     
		
        
        * Regression Model 2
        reg `var' cog interpers grit `controls' i.degreeschool if DGradColl`f'==1 & `var'>-9999
		matrix results = r(table)
        matrix a = results[1,1..3]'
        matrix b = results[4,1..3]' 
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for cognitive
		local cell_reference_cog = char(67) + string(`row_number')
		* Create the Excel cell reference for interpersonal
		local cell_reference_int = char(71) + string(`row_number')
		* Create the Excel cell reference for grit
		local cell_reference_grit = char(75) + string(`row_number')
		* Output the b coefficient and p-value 
		matrix cog_baseline = matrix(results[1,1..2]')
		matrix int_baseline = matrix(results[2,1..2]')
		matrix grit_baseline = matrix(results[3,1..2]')
		putexcel `cell_reference_cog' = matrix(cog_baseline)
		putexcel `cell_reference_int' = matrix(int_baseline)
		putexcel `cell_reference_grit' = matrix(grit_baseline)
        local rb = char(67) + string(45 + `f' - 1)
        putexcel C42 = ("College FEs") `rb' = (e(r2_a))
        
        * Regression Model 3
        reg `var' cog interpers grit `controls' i.degreeschool i.degreefield if DGradColl`f'==1 & `var'>-9999
       matrix results = r(table)
        matrix a = results[1,1..3]'
        matrix b = results[4,1..3]' 
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for cognitive
		local cell_reference_cog = char(68) + string(`row_number')
		* Create the Excel cell reference for interpersonal
		local cell_reference_int = char(72) + string(`row_number')
		* Create the Excel cell reference for grit
		local cell_reference_grit = char(76) + string(`row_number')
		* Output the b coefficient and p-value 
		matrix cog_baseline = matrix(results[1,1..2]')
		matrix int_baseline = matrix(results[2,1..2]')
		matrix grit_baseline = matrix(results[3,1..2]')
		putexcel `cell_reference_cog' = matrix(cog_baseline)
		putexcel `cell_reference_int' = matrix(int_baseline)
		putexcel `cell_reference_grit' = matrix(grit_baseline)
        local rc = char(68) + string(45 + `f' - 1)
        putexcel D42 = ("Program FEs") `rc' = (e(r2_a))
       
    
        
        * Regression Model 4
        reg `var' cog interpers grit `controls' i.degreefield if DGradColl`f'==1 & `var'>-9999 
		matrix results = r(table)
        matrix a = results[1,1..3]'
        matrix b = results[4,1..3]' 
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for cognitive
		local cell_reference_cog = char(69) + string(`row_number')
		* Create the Excel cell reference for interpersonal
		local cell_reference_int = char(73) + string(`row_number')
		* Create the Excel cell reference for grit
		local cell_reference_grit = char(77) + string(`row_number')
		* Output the b coefficient and p-value 
		matrix cog_baseline = matrix(results[1,1..2]')
		matrix int_baseline = matrix(results[2,1..2]')
		matrix grit_baseline = matrix(results[3,1..2]')
		putexcel `cell_reference_cog' = matrix(cog_baseline)
		putexcel `cell_reference_int' = matrix(int_baseline)
		putexcel `cell_reference_grit' = matrix(grit_baseline)
        local rd = char(69) + string(45 + `f' - 1)
        putexcel  E42 = ("Program FEs") `rd' = (e(r2_a))
        }
}



**COEFFICIENTS Over High School Tracks 

local controls "cog interpers grit fam_inc73_mom college_mom hs_mom hs_mom_miss college_dad hs_dad hs_dad_miss health1 health2 health_missing adveng advmath"
*Removed HS_track3 and HS_track4 from controls, added cog, interpers, grit

foreach var of varlist logwage logpvdispinc {
        * Set the excel output file and output sheet
        putexcel set college_major_program_school_FE_raw_v1_robustness_fields_clusters, modify sheet("Table_FE_`var'")
		*putexcel set "C:\Users\yuhangli\OneDrive - The University of Chicago\Documents\Summer 2023 SISRM RA\Update Tables\robustness_fields_clusters", modify sheet("Table_FE_`var'")
		
        * Outputting Results
        forvalues f=1/12 {
        sum DGradColl`f' 
        sum cog interpers grit HS_track3 HS_track4 if DGradColl`f'==1
        * Output Total Observation Number N
        local number1 = char(86) + string(7 + 2 * (`f' - 1))
        putexcel `number1' = (r(N))     
        * Set the header and legends
        putexcel V4 = ("N students") 
		***A46 = ("N with wage > 0") A47 = ("N with  wage > 0 and all FEs")
        
        *** UPDATED global headers for better table format
        **vertical headers 
                putexcel A6 = ("3-year")
                putexcel A7 = ("~~~(Hum, Soc Sci)") A9 = ("~~~(Sci, Math, Eng)") A11 = ("~~~Business")  A13 = ("Health Sci") A15 = ("Education") 
              
                putexcel A17 = ("~~~Humanities") A19 = ("~~~SocialSci")
        putexcel A21 = ("~~~Sciences") A23 = ("~~~Engineering") A25 = ("~~~Medicine")
                putexcel A27 = ("~~~Business") A29 = ("~~~Law")
                putexcel A4 = ("College Major")
                
                *Horizontal Headers
                putexcel B2 = ("Ability") N2 = ("Academic High School Track")
                putexcel B3 = ("Cognitive") F3 = ("Interpersonal") J3 = ("Grit")
                putexcel B4 = ("Baseline")  C4 = ("CollegeFEs")  D4 = ("ProgramFEs") E4 = ("College + ProgramFEs")
				
                putexcel F4 = ("Baseline")  G4 = ("CollegeFEs")  H4 = ("ProgramFEs") I4 = ("College + ProgramFEs")
				
               putexcel J4 = ("Baseline")  K4 = ("CollegeFEs")  L4 = ("ProgramFEs") M4 = ("College + ProgramFEs")
                               
                putexcel B40 = ("Adjusted R^2")
		
				putexcel N3 = ("nonSTEM") R3 = ("STEM")
				putexcel N4 = ("Baseline")  O4 = ("CollegeFEs")  P4 = ("ProgramFEs") Q4 = ("College + ProgramFEs")
				putexcel R4 = ("Baseline")  S4 = ("CollegeFEs")  T4 = ("ProgramFEs") U4 = ("College + ProgramFEs")
				putexcel B41 = ("Ability")  F41 = ("Academic High School Track")
				putexcel A16 = ("4 year")
         putexcel A3:X3, bold overwritefmt
        putexcel A4:X4, bold overwritefmt
        
        * Regression Model 1
		*reg logwage HS_track3 HS_track4 `controls' if DGradColl1==1 & logwage >-9999
        reg `var' HS_track3 HS_track4 `controls' if DGradColl`f'==1 & `var'>-9999
        matrix results = r(table)
        matrix a = results[1,1..2]'
        matrix b = results[4,1..2]'
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for HS_track3
		local cell_reference_HS_track3 = char(78) + string(`row_number')
		* Create the Excel cell reference for HS_track4
		local cell_reference_HS_track4 = char(82) + string(`row_number')
		* Output the b coefficient and p-value for HS_track3 and HS_track4
		matrix track_3 = matrix(results[1,1..2]')
		matrix track_4 = matrix(results[2,1..2]')
		putexcel `cell_reference_HS_track3' = matrix(track_3)
		putexcel `cell_reference_HS_track4' = matrix(track_4) 
        local ra = char(70) + string(45 + `f' - 1)
        putexcel F42 = ("Baseline")  `ra' = (e(r2_a))
        * Output Observation Number
        local number2 = char(87) + string(7 + 2 * (`f' - 1))
        putexcel W4 = ("N with wage > 0") `number2' = (e(N)) 
		

  
        * Regression Model 2
        reg `var' HS_track3 HS_track4 `controls' i.degreeschool if DGradColl`f'==1 & `var'>-9999
        matrix results = r(table)
        matrix a = results[1,1..2]'
        matrix b = results[4,1..2]'
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for HS_track3
		local cell_reference_HS_track3 = char(79) + string(`row_number')
		* Create the Excel cell reference for HS_track4
		local cell_reference_HS_track4 = char(83) + string(`row_number')
		* Output the b coefficient and p-value for HS_track3 and HS_track4
		matrix track_3 = matrix(results[1,1..2]')
		matrix track_4 = matrix(results[2,1..2]')
		putexcel `cell_reference_HS_track3' = matrix(track_3)
		putexcel `cell_reference_HS_track4' = matrix(track_4) 
        local rb = char(71) + string(45 + `f' - 1)
        putexcel G42 = ("College FEs") `rb' = (e(r2_a))
        
        * Regression Model 3
        reg `var' HS_track3 HS_track4 `controls' i.degreeschool i.degreefield if DGradColl`f'==1 & `var'>-9999
        matrix results = r(table)
        matrix a = results[1,1..2]'
        matrix b = results[4,1..2]'
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for HS_track3
		local cell_reference_HS_track3 = char(80) + string(`row_number')
		* Create the Excel cell reference for HS_track4
		local cell_reference_HS_track4 = char(84) + string(`row_number')
		* Output the b coefficient and p-value for HS_track3 and HS_track4
		matrix track_3 = matrix(results[1,1..2]')
		matrix track_4 = matrix(results[2,1..2]')
		putexcel `cell_reference_HS_track3' = matrix(track_3)
		putexcel `cell_reference_HS_track4' = matrix(track_4)
        local rc = char(72) + string(45 + `f' - 1)
        putexcel H42 = ("Program FEs") `rc' = (e(r2_a))
        * Output Observation Number
        local number3 = char(88) + string(7 + 2*(`f' - 1))
        putexcel X4 = ("N with  wage > 0 and all FEs") `number3' = (e(N))
        
        * Regression Model 4
        reg `var' HS_track3 HS_track4 `controls' i.degreefield if DGradColl`f'==1 & `var'>-9999
        matrix results = r(table)
        matrix a = results[1,1..2]'
        matrix b = results[4,1..2]'
        matrix results = a,b
	   * Calculate the row number based on the parity of f
        local row_number = 7 + 2 * (`f' - 1)
		* Create the Excel cell reference for HS_track3
		local cell_reference_HS_track3 = char(81) + string(`row_number')
		* Create the Excel cell reference for HS_track4
		local cell_reference_HS_track4 = char(85) + string(`row_number')
		* Output the b coefficient and p-value for HS_track3 and HS_track4
		matrix track_3 = matrix(results[1,1..2]')
		matrix track_4 = matrix(results[2,1..2]')
		putexcel `cell_reference_HS_track3' = matrix(track_3)
		putexcel `cell_reference_HS_track4' = matrix(track_4)
        local rd = char(73) + string(45 + `f' - 1)
        putexcel  I42 = ("Program FEs") `rd' = (e(r2_a))
        }
}

