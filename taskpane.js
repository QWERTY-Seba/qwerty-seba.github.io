var nombre_hoja = "Completar"
var columnas_hoja = null
var letra_primera_columna_encabezado = "A"
var letra_ultima_columna_encabezado = "K" //Columna con el dato de respuesta
var indice_columna_inspeccionar = 9 //respuesta
var letra_columna_inspeccionar = "J" //respuesta
var indice_primera_fila_encabezado = 1
 var delimitador_csv = ";"

var CANTIDAD_FILAS_CARGAR_POR_RONDA = 30 // filas x peticion
var MAX_NUMERO_FILAS_CONSULTAR = 100

async function traer_casos_no_procesados_ultimoRango(ultimo_indice){
    let resultado = []
    await Excel.run(async (context) => {
        let hoja = context.workbook.worksheets.getItem(nombre_hoja);
        
        let rango_encabezados = hoja.getRange(`${letra_primera_columna_encabezado}${indice_primera_fila_encabezado}:${letra_ultima_columna_encabezado}${indice_primera_fila_encabezado}`)
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
        resultado.push(["indice_excel"].concat(rango_encabezados.values[0]))      
        
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

            for(var [indice_,fila_] of valores_rango_cargado){
                if(fila_[0] == ""){
                    ultimo_indice_vacio_encontrado = true 
                    break;
                }
                let indice_fila = indice_limite_superior + indice_
                let rango_fila_columna_carga = `${letra_ultima_columna_encabezado}${indice_fila}` // A1
                resultado.push([rango_fila_columna_carga].concat(fila_))
            }
            
            indice_limite_superior = indice_limite_inferior
        }
    });

    if(resultado.length == 1){
        console.info(`No se encontraron casos nuevos entre ${letra_primera_columna_encabezado}${ultimo_indice}:${letra_ultima_columna_encabezado}${indice_limite_inferior} `)
        return []; 
    }
    
    return resultado;
}

function resultado_a_csv(resultado){
    return resultado.map(e => e.join(delimitador_csv)).join("\r")
}

async function pegar_csv(texto_csv){
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem(nombre_hoja);
        let valores_fila = texto_csv.split('\n').slice(1) //Quitar los encabezados
        for(let fila of valores_fila){
            fila = fila.split(delimitador_csv)
            let range = sheet.getRange(fila[0])
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
});




