const xlsx = require("xlsx");
const fs = require("fs");
///Importando datos del archivo xlsx y parseando a un array de objetos
const workbook1 = xlsx.readFile("prueba1.xlsx");
const sheetName1 = workbook1.SheetNames[0];
const worksheet1 = workbook1.Sheets[sheetName1];
const data1 = xlsx.utils.sheet_to_json(worksheet1);

const workbook2 = xlsx.readFile("prueba2.xlsx");
const sheetName2 = workbook2.SheetNames[0];
const worksheet2 = workbook2.Sheets[sheetName2];
const data2 = xlsx.utils.sheet_to_json(worksheet2);

///Función para manejar datos
function cantidadInventario(archivoWordpress, archivoODOO) {
  arrayJuntos = [];
  archivoWordpress.forEach((element) => {
    let sku = element?.SKU;
    archivoODOO.forEach((dataODOO, index) => {
      if (sku === dataODOO.SKU) {
        arrayJuntos.push({
          SKU: sku,
          Inventario_Wordpress: element.Inventario ? element.Inventario : "NA",
          Inventario_ODOO:
            dataODOO["Inventario"] >= 0
              ? dataODOO["Inventario"]
              : "No existe en ODOO",
        });
        archivoWordpress.splice(index, 1);
        archivoODOO.splice(index, 1);
      }
    });
  });
  return [...arrayJuntos, {SKU: "", Inventario_Wordpress: "", Inventario_ODOO: " "}, {SKU: "Los que no coinciden con ninguno de ODOO", Inventario_Wordpress: "", Inventario_ODOO: " "}, ...archivoODOO, {SKU: "", Inventario_Wordpress: "", Inventario_ODOO: " "}, {SKU: "Los que no coinciden con ninguno de Wordpress", Inventario_Wordpress: "", Inventario_ODOO: " "}, ...archivoWordpress];
}


///Lo de abajo exporta a un archivo de excel (xlsx)
const worksheet = xlsx.utils.json_to_sheet(cantidadInventario(data1, data2));

const workbook = xlsx.utils.book_new();

xlsx.utils.book_append_sheet(workbook, worksheet, "Inventario");

xlsx.writeFile(workbook, "Inventario.xlsx");
