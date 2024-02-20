var delimitador_csv = ";"
var CANTIDAD_FILAS_CARGAR_POR_RONDA = 30 // filas x peticion
var MAX_NUMERO_FILAS_CONSULTAR = 100 // Colocar valor mayor, como 1000, el objetivo es evitar problemas si es que selecciono un rango equivocado
var cache_ultimo_valores_consultados;
var hoja_actual;
var nombre_hoja = "Completar"

//USAR ESTA VARIABLE PARA ALMACENAR LOS VALORES QUE SE COLOQUEN EN EL HTML, HAY QUE VER COMO PERMITIR EL GUARDAR SOLO UN POCO
var hojas = {
}


function crear_interfaz_tabla(id_tabla = null, nombre_tabla = null, ultimo_indice = null, rango_encabezados = null){
    let html = `
    <h3>${nombre_tabla}</h3>
    <ul>
        <li>ultimo indice <span class="ultimo_indice">${ultimo_indice}</span> 
            <div class="contenedor_botones">    
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','ultimo_indice','buscar')"><p><i data-icon-name="BorderDrawToolGrid_20" aria-hidden="true" class="icon-339" style=" display: inline-block; width: 16px; "><svg height="100%" width="100%" viewBox="0,0,2048,2048" focusable="false"><path type="path" class="OfficeIconColors_HighContrast" d="M 1988 674 q 30 30 45 68 q 15 38 15 77 q 0 40 -15 77 q -15 38 -45 68 l -917 910 l -454 167 l 183 -468 l 898 -899 q 30 -30 68 -45 q 38 -14 77 -14 q 39 0 77 15 q 38 15 68 44 m -1005 1124 q -43 -76 -119 -119 l -71 188 m 1034 -887 l -144 -145 l -760 760 q 92 54 145 144 m 848 -847 q 15 -15 23 -34 q 7 -18 7 -38 q 0 -21 -8 -40 q -8 -18 -22 -33 q -15 -14 -33 -22 q -19 -8 -40 -8 q -20 0 -38 7 q -19 8 -34 23 l -16 16 l 145 144 m -1695 834 h 365 l -40 102 h -428 v -1741 h 1741 v 359 q -26 0 -51 4 q -26 5 -51 12 v -272 h -717 v 717 h 210 l -102 102 h -108 v 108 l -102 102 v -210 h -717 m 0 -819 v 717 h 717 v -717 z" style=" fill: transparent; "></path><path type="path" class="OfficeIconColors_m20" d="M 819 1634 q 240 -270 448 -482 q 59 -60 119 -120 q 60 -59 118 -113 q 57 -53 111 -99 q 53 -46 99 -80 q 46 -33 83 -52 q 36 -19 60 -19 q 8 0 12 1 q 41 11 68 28 q 26 17 41 37 q 15 20 21 41 q 6 22 6 42 q 0 17 -3 32 q -3 16 -7 27 q -4 13 -9 25 l -958 950 m -478 -60 h -396 v -1638 h 1638 v 312 q -58 8 -109 33 q -51 25 -93 67 l -921 921 z" style=" fill: #FAFAFAFF; "></path><path type="path" class="OfficeIconColors_m23" d="M 922 1024 h -768 v -102 h 768 v -768 h 102 v 768 h 210 l -102 102 h -108 v 108 l -102 102 z" style=" fill: #797774FF; "></path><path type="path" class="OfficeIconColors_m22" d="M 530 1843 h -428 v -1741 h 1741 v 359 q -26 0 -51 4 q -26 5 -51 12 v -272 h -1536 v 1536 h 365 z" style=" fill: #3A3A38FF; "></path><path type="path" class="OfficeIconColors_m26" d="M 836 1609 q 77 27 134 84 q 56 57 83 133 l -362 145 z"></path><path type="path" class="OfficeIconColors_m24" d="M 1988 674 q 30 30 45 68 q 15 38 15 77 q 0 40 -15 77 q -15 38 -45 68 l -917 910 l -454 167 l 183 -468 l 898 -899 q 30 -30 68 -45 q 38 -14 77 -14 q 39 0 77 15 q 38 15 68 44 m -1065 921 q 92 54 145 144 l 759 -759 l -144 -145 m -890 1032 l 190 -69 q -43 -76 -119 -119 m 1052 -787 q 15 -15 23 -34 q 7 -18 7 -38 q 0 -21 -8 -40 q -8 -18 -22 -33 q -15 -14 -33 -22 q -19 -8 -40 -8 q -20 0 -38 7 q -19 8 -34 23 l -16 16 l 145 144 z" style=" fill: #1E8BCDFF; "></path></svg></i>
                </p><span>Destacar</span></button>
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','ultimo_indice','editar')"><p>&#9999;&#65039;</p><span>Editar</span></button>
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','ultimo_indice','guardar')"><p>&#128190;</p><span>Guardar</span></button>
            </div>
        </li>
    
        <li>rango encabezados <span class="rango_encabezados">${rango_encabezados}</span>
            <div class="contenedor_botones">
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','rango_encabezados','buscar')"><p><i data-icon-name="BorderDrawToolGrid_20" aria-hidden="true" class="icon-339" style=" display: inline-block; width: 16px; "><svg height="100%" width="100%" viewBox="0,0,2048,2048" focusable="false"><path type="path" class="OfficeIconColors_HighContrast" d="M 1988 674 q 30 30 45 68 q 15 38 15 77 q 0 40 -15 77 q -15 38 -45 68 l -917 910 l -454 167 l 183 -468 l 898 -899 q 30 -30 68 -45 q 38 -14 77 -14 q 39 0 77 15 q 38 15 68 44 m -1005 1124 q -43 -76 -119 -119 l -71 188 m 1034 -887 l -144 -145 l -760 760 q 92 54 145 144 m 848 -847 q 15 -15 23 -34 q 7 -18 7 -38 q 0 -21 -8 -40 q -8 -18 -22 -33 q -15 -14 -33 -22 q -19 -8 -40 -8 q -20 0 -38 7 q -19 8 -34 23 l -16 16 l 145 144 m -1695 834 h 365 l -40 102 h -428 v -1741 h 1741 v 359 q -26 0 -51 4 q -26 5 -51 12 v -272 h -717 v 717 h 210 l -102 102 h -108 v 108 l -102 102 v -210 h -717 m 0 -819 v 717 h 717 v -717 z" style=" fill: transparent; "></path><path type="path" class="OfficeIconColors_m20" d="M 819 1634 q 240 -270 448 -482 q 59 -60 119 -120 q 60 -59 118 -113 q 57 -53 111 -99 q 53 -46 99 -80 q 46 -33 83 -52 q 36 -19 60 -19 q 8 0 12 1 q 41 11 68 28 q 26 17 41 37 q 15 20 21 41 q 6 22 6 42 q 0 17 -3 32 q -3 16 -7 27 q -4 13 -9 25 l -958 950 m -478 -60 h -396 v -1638 h 1638 v 312 q -58 8 -109 33 q -51 25 -93 67 l -921 921 z" style=" fill: #FAFAFAFF; "></path><path type="path" class="OfficeIconColors_m23" d="M 922 1024 h -768 v -102 h 768 v -768 h 102 v 768 h 210 l -102 102 h -108 v 108 l -102 102 z" style=" fill: #797774FF; "></path><path type="path" class="OfficeIconColors_m22" d="M 530 1843 h -428 v -1741 h 1741 v 359 q -26 0 -51 4 q -26 5 -51 12 v -272 h -1536 v 1536 h 365 z" style=" fill: #3A3A38FF; "></path><path type="path" class="OfficeIconColors_m26" d="M 836 1609 q 77 27 134 84 q 56 57 83 133 l -362 145 z"></path><path type="path" class="OfficeIconColors_m24" d="M 1988 674 q 30 30 45 68 q 15 38 15 77 q 0 40 -15 77 q -15 38 -45 68 l -917 910 l -454 167 l 183 -468 l 898 -899 q 30 -30 68 -45 q 38 -14 77 -14 q 39 0 77 15 q 38 15 68 44 m -1065 921 q 92 54 145 144 l 759 -759 l -144 -145 m -890 1032 l 190 -69 q -43 -76 -119 -119 m 1052 -787 q 15 -15 23 -34 q 7 -18 7 -38 q 0 -21 -8 -40 q -8 -18 -22 -33 q -15 -14 -33 -22 q -19 -8 -40 -8 q -20 0 -38 7 q -19 8 -34 23 l -16 16 l 145 144 z" style=" fill: #1E8BCDFF; "></path></svg></i>
                </p><span>Destacar</span></button>
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','rango_encabezados','editar')"><p>&#9999;&#65039;</p><span>Editar</span></button>
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','rango_encabezados','guardar')"><p>&#128190;</p><span>Guardar</span></button>
            </div>
        </li>
    
        <li>rango validacion <span class="rango_validacion"></span> 
            <div class="contenedor_botones">
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','ultimo_indice','buscar')"><p><i data-icon-name="BorderDrawToolGrid_20" aria-hidden="true" class="icon-339" style=" display: inline-block; width: 16px; "><svg height="100%" width="100%" viewBox="0,0,2048,2048" focusable="false"><path type="path" class="OfficeIconColors_HighContrast" d="M 1988 674 q 30 30 45 68 q 15 38 15 77 q 0 40 -15 77 q -15 38 -45 68 l -917 910 l -454 167 l 183 -468 l 898 -899 q 30 -30 68 -45 q 38 -14 77 -14 q 39 0 77 15 q 38 15 68 44 m -1005 1124 q -43 -76 -119 -119 l -71 188 m 1034 -887 l -144 -145 l -760 760 q 92 54 145 144 m 848 -847 q 15 -15 23 -34 q 7 -18 7 -38 q 0 -21 -8 -40 q -8 -18 -22 -33 q -15 -14 -33 -22 q -19 -8 -40 -8 q -20 0 -38 7 q -19 8 -34 23 l -16 16 l 145 144 m -1695 834 h 365 l -40 102 h -428 v -1741 h 1741 v 359 q -26 0 -51 4 q -26 5 -51 12 v -272 h -717 v 717 h 210 l -102 102 h -108 v 108 l -102 102 v -210 h -717 m 0 -819 v 717 h 717 v -717 z" style=" fill: transparent; "></path><path type="path" class="OfficeIconColors_m20" d="M 819 1634 q 240 -270 448 -482 q 59 -60 119 -120 q 60 -59 118 -113 q 57 -53 111 -99 q 53 -46 99 -80 q 46 -33 83 -52 q 36 -19 60 -19 q 8 0 12 1 q 41 11 68 28 q 26 17 41 37 q 15 20 21 41 q 6 22 6 42 q 0 17 -3 32 q -3 16 -7 27 q -4 13 -9 25 l -958 950 m -478 -60 h -396 v -1638 h 1638 v 312 q -58 8 -109 33 q -51 25 -93 67 l -921 921 z" style=" fill: #FAFAFAFF; "></path><path type="path" class="OfficeIconColors_m23" d="M 922 1024 h -768 v -102 h 768 v -768 h 102 v 768 h 210 l -102 102 h -108 v 108 l -102 102 z" style=" fill: #797774FF; "></path><path type="path" class="OfficeIconColors_m22" d="M 530 1843 h -428 v -1741 h 1741 v 359 q -26 0 -51 4 q -26 5 -51 12 v -272 h -1536 v 1536 h 365 z" style=" fill: #3A3A38FF; "></path><path type="path" class="OfficeIconColors_m26" d="M 836 1609 q 77 27 134 84 q 56 57 83 133 l -362 145 z"></path><path type="path" class="OfficeIconColors_m24" d="M 1988 674 q 30 30 45 68 q 15 38 15 77 q 0 40 -15 77 q -15 38 -45 68 l -917 910 l -454 167 l 183 -468 l 898 -899 q 30 -30 68 -45 q 38 -14 77 -14 q 39 0 77 15 q 38 15 68 44 m -1065 921 q 92 54 145 144 l 759 -759 l -144 -145 m -890 1032 l 190 -69 q -43 -76 -119 -119 m 1052 -787 q 15 -15 23 -34 q 7 -18 7 -38 q 0 -21 -8 -40 q -8 -18 -22 -33 q -15 -14 -33 -22 q -19 -8 -40 -8 q -20 0 -38 7 q -19 8 -34 23 l -16 16 l 145 144 z" style=" fill: #1E8BCDFF; "></path></svg></i>
                </p><span>Destacar</span></button>
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','ultimo_indice','editar')"><p>&#9999;&#65039;</p><span>Editar</span></button>
                <button class="boton_rangos" onclick="cargar_seleccion('${id_tabla}','ultimo_indice','guardar')"><p>&#128190;</p><span>Guardar</span></button>
            </div>
        </li>
        <button onclick="datos_a_descargar_prueba('${id_tabla}')">Extraccion de Prueba</button>
    </ul>`

    let div = document.createElement("div")
    div.id = id_tabla
    div.classList.add("tabla")
    div.innerHTML = html
    
    document.querySelector("#contenedor_tablas").appendChild(div)

}

function crear_tabla_personalizada(){

    let id_tabla = (Math.random() + 1).toString(36).substring(2)

    crear_interfaz_tabla(id_tabla = id_tabla, nombre_tabla = id_tabla)
    
    hojas[hoja_actual].tablas[id_tabla] = {
        ultimo_indice : null,
        nombre_tabla : id_tabla,
        rango_completo_tabla : null,
        rango_encabezados : null,
        valores_encabezados : [],
        personalizada : true
    }


}

//AGREGAR FUNCION PARA VERIFICAR SI EL ORDEN DE LOS ENCABEZADOS CAMBIO
async function extraer_nombres_hojas(){
    await Excel.run(async ctx => {
        let lista_hojas = ctx.workbook.worksheets
        let hoja_activa = ctx.workbook.worksheets.getActiveWorksheet();

        lista_hojas.load(["name","id"])
        await ctx.sync()        

        hoja_actual = hoja_activa.id

        for(let hoja of lista_hojas.items){
            hojas[hoja.id] = {
                nombre_hoja : hoja.name,
                tablas : {}
            }
            hoja.onActivated.add(e => {actualizar_panel(e)})
        }
        await ctx.sync()
    })
}

async function actualizar_panel(event){
    document.querySelector("#nombre_hoja").innerText = hojas[event.worksheetId].nombre_hoja
    hoja_actual = event.worksheetId 

}

//AL MANEJAR TABLAS, REVISAR SI EL RANGO CAMBIO

//CREAR FUNCION ACTUALIZAR PANEL, QUE TOME LA HOJA ACTIVA
//ACA ESTAMOS ASUMIENDO QUE LA ID DE TABLA ESTA EN LA HOJA ACTUAL, HABRIA QUE BUSCAR A LA HOJA QUE PERTENECE O MANDAR LA HOJA COMO PARAMETRO
async function datos_a_descargar_prueba(id_tabla){
    let temp_tabla = hojas[hoja_actual].tablas[id_tabla]

    if(temp_tabla == undefined){
        mostrar_error("id_tabla mal escrito o no existe")
        throw "id_tabla mas escrito o no existe"
    }
    if(temp_tabla.ultimo_indice == null || temp_tabla.rango_encabezados == null){
        mostrar_error(`uno o ambos valores de tabla nullos ultimo_indice: ${temp_tabla.ultimo_indice} rango_encabezados: ${temp_tabla.rango_encabezados}` )
        throw "Valores de tabla nullos"
    }


    cache_ultimo_valores_consultados = await traer_casos_no_procesados_ultimoRango(temp_tabla.ultimo_indice, temp_tabla.rango_encabezados)
    
    if(cache_ultimo_valores_consultados.length == 0){
        mostrar_error("No hay registros nuevos")
    }
    let indice_final_datos = cache_ultimo_valores_consultados.slice(-1)[0][0]

    let indice_descargar_datos = `${temp_tabla.ultimo_indice}:${indice_final_datos}`
    
    destacar_celdas(indice_descargar_datos)
}

function validar_superposicion_rango_tabla( ){

}

async function extraer_tablas_en_hoja(nombre_hoja){

    let resp = await Excel.run(async ctx => {
        let tablas = ctx.workbook.worksheets.getItem(nombre_hoja).tables.load("items")
        await ctx.sync()
        let lista_tablas = []
        for(var tabla of tablas.items){

            let rango = tabla.getRange()
            rango.load("address") 
            //TODAS LAS TABLAS DEFINIDAS TIENEN RANGO DE ENCABEZADOS

            let estructura_tabla = {
                id_tabla : tabla.id,
                ultimo_indice : null,
                nombre_tabla : tabla.name,
                rango_completo_tabla : rango,
                rango_encabezados : null,
                valores_encabezados : [],
                personalizada : false
            }

            lista_tablas.push(estructura_tabla)
        }
        await ctx.sync()
        
        lista_tablas.forEach((tabla, index) => {
            lista_tablas[index].rango_completo_tabla = tabla.rango_completo_tabla.address
            let rango_encabezados = extraer_rango(tabla.rango_completo_tabla)
            lista_tablas[index].rango_encabezados = `${rango_encabezados[1]}${rango_encabezados[2]}:${rango_encabezados[3]}${rango_encabezados[2]}`
            //BUSCAR ACA EL ULTIMO INDICE

            hojas[hoja_actual].tablas[tabla.id_tabla] = tabla
        })
    })
}

//SI NO EXISTE CONFIG, BUSCAR EN LOCAL
async function cargar_config(){
    //REVISAR SI EXISTEN LAS KEYS EN SETTINGS, COLOCAR ADVERTENCIA | CARGAR
    Office.context.document.settings.refreshAsync(async function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            rango_encabezados = Office.context.document.settings.get('rango_encabezados')
           
            //VALIDAR LA CONFIG CON LOS VALORES ACTUALES
            traer_rangos_tabla_en_hoja(hoja_actual)

            
        } else {
            mensaje = 'Error reading settings'
        }
    });

}
//TALVEZ HACER QUE ESTA FUNCION LANCE EL ERROR
function mostrar_error(mensaje){
    panel = document.querySelector('#errores')
    
    let div = document.createElement('div')
    div.classList.add("div_error")
    div.innerText = mensaje


    boton_cerrar = document.createElement('button')
    boton_cerrar.classList.add("boton_error")
    //b_p = document.createElement('p')
    boton_cerrar.innerHTML = "&#10005;"
    //boton_cerrar.appendChild(b_p)
    boton_cerrar.addEventListener("click", function(){this.parentElement.remove()})

    
    div.appendChild(boton_cerrar)
    panel.appendChild(div)
    
}

function extraer_rango(rango){
    regex = new RegExp(/(?<=!)([A-Z]+)(\d+)(?:\:([A-Z]+)(\d+))?/)
    resultado = regex.exec(rango)
    // if(resultado.length < 5){
    //     mostrar_error("Error extraer valores de rango")
    //     throw "Error extraer valores de rango"
    // } 
    return resultado
}

//CAMBIAR DE COLOR TEMPORALMENTE A LAS CELDAS QUE SE VAN A DESCARGAR
async function destacar_celdas(rango){
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange(rango);
        range.select();
        await context.sync();
    });
    

}

function calcular_indice_rango(rango_excel){
    var alpabeto = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    var largo = alpabeto.length
    var rango_transformado = extraer_rango(rango_excel)[1]
    //EL RANGO TIENE QUE LLEGAR COMO "A1"        
    var acumulado = 0
    for(var [i,valor] of rango_transformado.split("").reverse().entries()){
        acumulado +=  (alpabeto.indexOf(valor) + 1) *  (largo ** i ) 
    }    
    return acumulado
}

async function extraer_rango_seleccion(){
    var resp;
    await Excel.run(async function (ctx) {
        let selectedRange = ctx.workbook.getSelectedRange();
        selectedRange.load(["address","values"])
        
        try{
            await ctx.sync()
        }catch(Error){
            if(Error.message == 'The current selection is invalid for this operation.'){
                mostrar_error("realiza solo 1 seleccion")
                throw "realiza solo 1 seleccion"
                }
        }
        resp = selectedRange
    })
    return resp;
}


//FALTA ENVIAR MENSAJE DE COMO SALIO LA OPERACION
//CAMBIAR DE ID A ATRIBUTO EN EL DIV DE LA TABLA
async function cargar_seleccion(id_tabla, tipo, modo){        
    if(modo == "buscar"){
        
        if(hojas[hoja_actual].tablas[id_tabla][tipo] == null){
            mostrar_error(`valor nullo ${tipo}`)
            throw "Valores de tabla nullos"
        }

        destacar_celdas(hojas[hoja_actual].tablas[id_tabla][tipo])

        return;
    }
    
    
    let selectedRange = await extraer_rango_seleccion()
    
    switch(tipo){
        case "ultimo_indice":
            if(modo == "editar"){
                if(selectedRange.values.length > 1){
                    mostrar_error("mas de 1 fila seleccionada")
                    throw "mas de 1 fila seleccionada"
                }
                if(selectedRange.values[0].length > 1){
                    mostrar_error("mas de 1 columna seleccionada")
                    throw "mas de 1 columna seleccionada"
                }
                hojas[hoja_actual].tablas[id_tabla].ultimo_indice = selectedRange.address
            }
            
            if(modo == "guardar"){
                // Office.context.document.settings.set('rango_ultimo_indice', rango_ultimo_indice); 
                // Office.context.document.settings.saveAsync();
            }

            break
        case "rango_encabezados":
            if(modo == "editar"){        
                if(selectedRange.values.length > 1){
                        mostrar_error("mas de 1 fila seleccionada")
                        throw "mas de 1 fila seleccionada"
                    }
                    if(selectedRange.values[0].length < 2){
                        mostrar_error("seleccionar al menos 2 columas")
                        throw "seleccionar al menos 2 columas"
                    }
                    hojas[hoja_actual].tablas[id_tabla].rango_encabezados = selectedRange.address
                }
            if(modo == "guardar"){
                // Office.context.document.settings.set('rango_encabezados', rango_encabezados); 
                // Office.context.document.settings.saveAsync();
            }
            
            break
            //pegar_encabezados(selectedRange)
        case "rango_validacion":
            break
    } 

    document.querySelector(`div[id="${id_tabla}"] span.${tipo}`).innerText = selectedRange.address
}

//DEVOLVER O ACTUALIZAR EL ULTIMO INDICE
/*
Esto deberia funcionar con los parametros, ultimo rango y rango encabezados, crear una funcion aparte para llamar esta funcion con los parametros de la hoja o como sea

*/

/**
 * 
 * @param {string} p_ultimo_indice con cualquier formato de address de excel
 * @param {string} p_rango_encabezados con cualquier formato de address de excel
 * @returns 
 */

async function traer_casos_no_procesados_ultimoRango(p_ultimo_indice, p_rango_encabezados){
        
    let letra_primera_columna_encabezado, indice_primera_fila_encabezado , letra_ultima_columna_encabezado;
    [letra_primera_columna_encabezado, indice_primera_fila_encabezado, letra_ultima_columna_encabezado] = extraer_rango(p_rango_encabezados).slice(1)
    
    //TALVEZ HACER QUE LA FUNCION EXTRAER_RANGO DEVUELVA UN OBJETO EN VEZ DE LISTA, PARA QUE SEA MAS ENTENDIBLE LA DESENPAQUETACION
    let ultimo_indice_rango = extraer_rango(p_ultimo_indice)
    let ultimo_indice = ultimo_indice_rango[2]
    let columna_inspeccionar = calcular_indice_rango(p_ultimo_indice) - calcular_indice_rango(p_rango_encabezados) 
    
    if(columna_inspeccionar < 0 ){
        mostrar_error("error al calcular columna de inspeccionar")
        throw "error al calcular columna de inspeccionar"
    } 
    let valores_rango = []
    
    await Excel.run(async (context) => {
        let hoja = context.workbook.worksheets.getItem(nombre_hoja);
        //FALTA VALIDAR QUE HAYA AL MENOS 2 COLUMNAS, SOLO 1 DA PROBLEMAS
        let valores_encabezados = hoja.getRange(`${letra_primera_columna_encabezado}${indice_primera_fila_encabezado}:${letra_ultima_columna_encabezado}${indice_primera_fila_encabezado}`)
                                    .getUsedRange()
                                    .load("values")
        let mensaje = ""
        try{
            await context.sync();
        }catch(Error){
            switch(Error.code){
                case "ItemNotFound":
                    mensaje = `La hoja a cargar no existe en el archivo de Excel`
                    break
                case "InvalidArgument":
                    mensaje = `Uno de los parametros del rango o es valido o no se encuentra`
                    break
                default:
                    mensaje = `Error no controlado durante carga de encabezados`
            }
            mostrar_error(mensaje)
            throw Error;
        }
        valores_rango.push(["indice_excel"].concat(valores_encabezados.values[0]))      
        
        let ultimo_indice_vacio_encontrado = false
        indice_limite_superior = Number(ultimo_indice)
        
        let cantidad_filas_cargadas = 0
        
        while(ultimo_indice_vacio_encontrado == false){
            cantidad_filas_cargadas += CANTIDAD_FILAS_CARGAR_POR_RONDA
            
            if(cantidad_filas_cargadas >= MAX_NUMERO_FILAS_CONSULTAR){
                throw new Error("Se ha exedido el maximo de filas a consultar")
            }
            
            indice_limite_inferior = indice_limite_superior + CANTIDAD_FILAS_CARGAR_POR_RONDA 
            
            let rango_consultar = `${letra_primera_columna_encabezado}${indice_limite_superior}:${letra_ultima_columna_encabezado}${indice_limite_inferior}`
            
            let valores_rango_precarga = hoja.getRange(rango_consultar).load("values")
            await context.sync()

            let valores_rango_cargado = valores_rango_precarga.values.entries()
            //EXTRAER SOLAMENTE EL RANGO O DEVOLVER EL OBJETO DE RANGO
            for(var [indice_,fila_] of valores_rango_cargado){
                //ACA TIENE QUE ESTAR LA COLUMNA A INSPECCIONAR QUE SEA DEL ULTIMO INDICE
                if(fila_[columna_inspeccionar] == ""){
                    ultimo_indice_vacio_encontrado = true 
                    break;
                }
                let indice_fila = indice_limite_superior + indice_
                let rango_fila_columna_carga = `${letra_ultima_columna_encabezado}${indice_fila}` // A1
                valores_rango.push([rango_fila_columna_carga].concat(fila_))
            }
            
            indice_limite_superior = indice_limite_inferior
        }
    });

    if(valores_rango.length == 1){
        console.info(`No se encontraron casos nuevos entre ${letra_primera_columna_encabezado}${ultimo_indice}:${letra_ultima_columna_encabezado}${indice_limite_inferior} `)
        return []; 
    }
    
    return valores_rango;
}

function resultado_a_csv(resultado){
    return resultado.map(e => e.join(delimitador_csv)).join("\r")
}

async function pegar_csv(texto_csv){
    await Excel.run(async (context) => {
        let hoja = context.workbook.worksheets.getItem(nombre_hoja);
        let valores_fila = texto_csv.split('\n').slice(1) //Quitar los encabezados
        for(let fila of valores_fila){
            fila = fila.split(delimitador_csv)
            let range = hoja.getRange(fila[0])
            range.values = fila[1]//Quitar el indice_excel 
        }
        
        await context.sync();
    });

}

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
        Office.context.document.settings.saveAsync();
        
    }
    cargar_config()
});




