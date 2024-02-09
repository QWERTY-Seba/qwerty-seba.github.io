var nombre_hoja = "Completar"

//CAMBIAR A RANGO ENCABEZADOS

var rango_ultimo_indice = null;
var rango_encabezados = null;
var columna_validacion = null;
var delimitador_csv = ";"

var CANTIDAD_FILAS_CARGAR_POR_RONDA = 30 // filas x peticion
var MAX_NUMERO_FILAS_CONSULTAR = 100 // Colocar valor mayor, como 1000, el objetivo es evitar problemas si es que selecciono un rango equivocado
var cache_ultimo_valores_consultados;

async function datos_a_descargar_prueba(){
    //EXTRAER RANGO DE ULTIMO INDICE Y 
    /*
    crear rango a partir del ultimo rango mas encabezado mas el buscar ultimo indice
    cargar el rango con excel y seleccionarlo
    almacenar el rango y cambiar el formato
    quitar el formato despues de un time sleep
    */ 
    cache_ultimo_valores_consultados = await traer_casos_no_procesados_ultimoRango(rango_ultimo_indice, rango_encabezados)
    let indice_final_datos = cache_ultimo_valores_consultados.slice(-1)[0][0]

    let indice_descargar_datos = `${letra_primera_columna_encabezado}${i}:${indice_final_datos}`
    Excel.run( async ctx => {
        ctx.workbook.worksheets.getActiveWorksheet()
                .getRange(indice_descargar_datos)
                .select()
        await ctx.sync()

    })
}

async function cargar_config(){
    //REVISAR SI EXISTEN LAS KEYS EN SETTINGS, COLOCAR ADVERTENCIA | CARGAR
    Office.context.document.settings.refreshAsync(async function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            rango_encabezados = Office.context.document.settings.get('rango_encabezados')
            var no_pasar = false
            //ULTIMO INDICE CONTIENE LETRA Y NUMERO
            rango_ultimo_indice = Office.context.document.settings.get('ultimo_indice')
            
            if(rango_encabezados == null || rango_encabezados == undefined){
                //CREAR DIV QUE DIGA QUE NO ESTAN
                document.querySelector('span#rango_encabezados').innerText = "no encontrado"
                no_pasar = true
            }

            if(rango_ultimo_indice == null || rango_ultimo_indice == undefined){
                //CREAR DIV QUE DIGA QUE NO ESTAN
                document.querySelector('span#ultimo_indice').innerText = "no encontrado"
                no_pasar = true
            }

            if(no_pasar){
                return;
            }
            //TALVEZ GUARDAR EL RANGO COMO LISTA Y LISTO
            [letra_primera_columna_encabezado, indice_primera_fila_encabezado, letra_ultima_columna_encabezado] = extraer_rango(rango_encabezados).slice(1)

            rango_y_valores = await Excel.run(async function (ctx) {
                let hoja = context.workbook.worksheets.getItem(nombre_hoja);
                let selectedRange = hoja.getRange(`${letra_primera_columna_encabezado}${indice_primera_fila_encabezado}:${letra_ultima_columna_encabezado}${indice_primera_fila_encabezado}`);
                selectedRange.load(["address","values"])
                await ctx.sync()
                return selectedRange
            });

            pegar_encabezados(rango_y_valores)
            console.log(mySettings);
        } else {
            console.error('Error reading settings');
        }
    });

}
//CREAR OBJECTO QUE SE PUEDA QUITAR Y QUE VAYA EN EL DIV DE LA EXTENSION
//TALVEZ HACER QUE ESTA FUNCION LANCE EL ERROR
function mostrar_error(mensaje){
    panel = document.querySelector('body')
    
    let div = document.createElement('div')
    div.classList.add("div_error")
    div.innerText = mensaje


    boton_cerrar = document.createElement('button')
    boton_cerrar.classList.add("boton_error")
    b_p = document.createElement('p')
    b_p.innerHTML = "&#10005;"
    boton_cerrar.addEventListener("click", function(){this.parentElement.remove()})

    boton_cerrar.appendChild(b_p)
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
    Excel.run((context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange(rango);
    
        range.select();
    
        context.sync();
    });
    

}

function calcular_indice_rango(rango_excel){
    
    var alpabeto = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    var largo = alpabeto.length
    var rango_transformado = extraer_rango(rango_excel)[1]
    //TRANSFORMAR CADA LETRA DE ALGO EN NUMERO
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


async function cargar_seleccion(tipo){        
    let selectedRange = await extraer_rango_seleccion()
    
    switch(tipo){
        case "ultimo_indice":
            if(selectedRange.values.length > 1){
                mostrar_error("mas de 1 fila seleccionada")
                throw "mas de 1 fila seleccionada"
            }
            if(selectedRange.values[0].length > 1){
                mostrar_error("mas de 1 columna seleccionada")
                throw "mas de 1 columna seleccionada"
            }
            rango_ultimo_indice = selectedRange.address
        case "rango_encabezados":
            if(selectedRange.values.length > 1){
                mostrar_error("mas de 1 fila seleccionada")
                throw "mas de 1 fila seleccionada"
            }
            rango_encabezados = selectedRange.address
            //pegar_encabezados(selectedRange)
        case "rango_validacion":
            
    } 
    document.querySelector("span#" + tipo).innerText = selectedRange.address
}

function pegar_encabezados(range_y_valores){
    let L_1, N_1;
    L_1, N_1 = extraer_rango(range_y_valores.address)

    div = document.querySelector('#divEncabezados>select')
        div.innerHTML = ""
        //EXTRAER LA ADDRESS DE CADA ELEMENTO
        for(var valor of range_y_valores.values){

            
            // let check_columna = document.createElement('input')
            // let label = document.createElement('label')

            // let id_columna = "x" + i
            // //AL PARECER ESTA ES LA UNICA FORMA DE ASIGNAR ATRIBUTOS
            // //AGREGAR OPCIONES PARA QUE CADA COLUMNA PUEDA VALIDAR DE CIERTA MANERA
            // check_columna.setAttribute("type", "checkbox")
            // check_columna.setAttribute("name", valor)
            // check_columna.setAttribute("value", valor)
            // check_columna.setAttribute("id", id_columna)

            // label.innerText =  valor
            // label.setAttribute("for", id_columna)
            let option = document.createElement('option')
            option.setAttribute("value", valor)
            option.setAttribute("id", L_1 + i)
            option.innerText = valor
            
            div.appendChild(option)
            
        }
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
        throw "error al calcular columna de inspeccionar"
    } 
    let valores_rango = []
    
    await Excel.run(async (context) => {
        let hoja = context.workbook.worksheets.getItem(nombre_hoja);
        
        let valores_encabezados = hoja.getRange(`${letra_primera_columna_encabezado}${indice_primera_fila_encabezado}:${letra_ultima_columna_encabezado}${indice_primera_fila_encabezado}`)
                                    .getUsedRange()
                                    .load("values")
        
        try{
            await context.sync();
        }catch(Error){
            switch(Error.code){
                case "ItemNotFound":
                    console.error(`La hoja a cargar no existe en el archivo de Excel`);
                    break
                case "InvalidArgument":
                    console.error(`Uno de los parametros del rango o es valido o no se encuentra`)
                    break
                default:
                    console.error(`Error no controlado durante carga de encabezados`)
            }
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




