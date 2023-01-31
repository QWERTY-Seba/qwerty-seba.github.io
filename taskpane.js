// function traer_crear_div_resultados(){
//     let id_div = "div_seba_manuel_id_unica"
//     let div = document.getElementById(id_div)
//     if(div == undefined){
//         div = document.createElement('div')
//         div.id = id_div
//         document.body.appendChild(div)
//         return div;
//     }
//     div.innerHTML = ""
//     return div;
// }

var hoja_actual = "Sheet1"

// function traer_indice_columna_por_nombre(hoja, nombre_columna){
//     return;
// }

//Probablemente la function necesite que retorne los indices de cada fila que se extraiga para tener referencia al momento de volver a insertar
async function traer_casos_no_procesados(){
    var cantidad_filas_ronda = 10
    var resultado = []

    await Excel.run(async (context) => {
        let hoja = context.workbook.worksheets.getItem(hoja_actual);

        //Cambiar a dinamico, para encontrar indice de columna necesario no escribirlo directamente
        let indice_ultimo_rango = hoja.getRange("A:A")
                                    .getUsedRange()
                                    .getLastRow()
                                    .load("address")
        await context.sync();
        let indice_superior = /[!][A-Z]+(\d+)/.exec(indice_ultimo_rango.address)[1]
        indice_superior = Number(indice_superior)
        let indice_inferior = 0
        
        let headers = []

        for(var i = 0; i<5; i++){
            indice_inferior = indice_superior - cantidad_filas_ronda 
            indice_inferior = indice_inferior < 1 ? 1 : indice_inferior //evitar que baje del minimo
            let indice_rangos = `${hoja_actual}!A${indice_superior}:E${indice_inferior}`
            console.log(indice_rangos)
            let rangos = hoja.getRange(indice_rangos).load("values")
            await context.sync()
            for(var fila of rangos.values.reverse()){
                if(fila[4] != ""){
                    indice_inferior = 1 //Probablemente no sea lo mas elegante
                    break;
                }
                resultado.push(fila)
            }
            if(indice_inferior === 1){
                break;
            }
            indice_superior = indice_inferior
        }
        // div = traer_crear_div_resultados()
        // div.innerText = JSON.stringify(resultado)
        
    });
    return resultado;
}

//Office.actions.associate('PASTECLIPBOARD', traer_casos_no_procesados());

Office.onReady((info) => {
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
        Office.context.document.settings.saveAsync();
    }
});




