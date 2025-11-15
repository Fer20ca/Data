forvalues year = 2007/2024 {
    
    display "Procesando el año `year'..."

    * Definir directorio de entrada
    cd "C:\Users\FERNANDO\Documents\PI_INEQUIDAD\scripts_data\data\inputs\sumaria_enaho\sumarias_dta"

    * Crear el nombre del archivo de entrada y abrirlo
    local input_filename = "sumaria-" + "`year'" + ".dta"
    use "`input_filename'", clear

    * Variables básicas
    gen gpcm = (gashog2d / (mieperho*12))
    gen ipcm = (inghog1d / (mieperho*12))
    gen fac2 = factor07 * mieperho  // Asegúrate que el nombre del factor cambie con el año si es necesario

    * Generar variable de departamento
    gen dpto = real(substr(ubigeo,1,2))
    replace dpto = 15 if dpto == 7
    label define dpto 1"Amazonas" 2"Ancash" 3"Apurimac" 4"Arequipa" 5"Ayacucho" 6"Cajamarca" 8"Cusco" ///
        9"Huancavelica" 10"Huanuco" 11"Ica" 12"Junin" 13"La Libertad" 14"Lambayeque" 15"Lima" 16"Loreto" ///
        17"Madre de Dios" 18"Moquegua" 19"Pasco" 20"Piura" 21"Puno" 22"San Martin" 23"Tacna" 24"Tumbes" 25"Ucayali"
    label values dpto dpto

    * Calcular desigualdad
    ineqdec0 ipcm [aw=fac2], byg(dpto)

    * Preparar variables para exportar
    gen str20 n_dep = ""
    replace n_dep = "Amazonas"       in 1
    replace n_dep = "Ancash"         in 2
    replace n_dep = "Apurimac"       in 3
    replace n_dep = "Arequipa"       in 4
    replace n_dep = "Ayacucho"       in 5
    replace n_dep = "Cajamarca"      in 6
    replace n_dep = "Callao"         in 7
    replace n_dep = "Cusco"          in 8
    replace n_dep = "Huancavelica"   in 9
    replace n_dep = "Huanuco"        in 10
    replace n_dep = "Ica"            in 11
    replace n_dep = "Junin"          in 12
    replace n_dep = "La Libertad"    in 13
    replace n_dep = "Lambayeque"     in 14
    replace n_dep = "Lima"           in 15
    replace n_dep = "Loreto"         in 16
    replace n_dep = "Madre de Dios"  in 17
    replace n_dep = "Moquegua"       in 18
    replace n_dep = "Pasco"          in 19
    replace n_dep = "Piura"          in 20
    replace n_dep = "Puno"           in 21
    replace n_dep = "San Martin"     in 22
    replace n_dep = "Tacna"          in 23
    replace n_dep = "Tumbes"         in 24
    replace n_dep = "Ucayali"        in 25

    gen gini = .
    gen theil = .
    gen gini_g = .
    gen theil_g = .
    gen ratio_g = .

    replace gini_g = r(gini)      in 1
    replace theil_g = r(ge2)      in 1
    replace ratio_g = r(p90p10)   in 1

    forvalues k = 1/25 {
        if `k' != 7 {
            quietly replace gini = r(gini_`k') if _n == `k'
            quietly replace theil = r(ge2_`k') if _n == `k'
        }
    }

    * Exportar resultados
    cd "C:\Users\FERNANDO\Documents\PI_INEQUIDAD\scripts_data\data\outputs\gini"
    local output_filename = "gini_" + "`year'" + ".xlsx"
    export excel n_dep gini theil gini_g theil_g ratio_g using "`output_filename'", firstrow(variables) replace

}
