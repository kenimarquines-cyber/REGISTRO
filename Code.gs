// Pegar en Extensiones > Apps Script de tu Google Sheet (reemplaza el código anterior).
const HOJA = "Registros";
const CATALOGO = "Catalogo";   // TABLA MADRE
const CARPETA_FOTOS = "Fotos instalacion";

// Devuelve las listas de la tabla madre para el autocompletado
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CATALOGO);
  if (!sh) {
    sh = ss.insertSheet(CATALOGO);
    sh.appendRow(["Nombre","Cargo","Punto de venta","Material"]);
    sh.appendRow(["Kenneth Pérez","Supervisor de tienda","Sucursal Centro","Banner"]);
    sh.appendRow(["Ana Gómez","Supervisor de zona","Sucursal Norte","Exhibidor"]);
    sh.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1f3a5f").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  const v = sh.getDataRange().getValues().slice(1);
  const uniq = a => [...new Set(a.map(x => String(x).trim()).filter(Boolean))];
  const nombres = v.filter(r => String(r[0]).trim())
                   .map(r => ({ n: String(r[0]).trim(), c: String(r[1]).trim() }));
  return salida({
    nombres: nombres,
    cargos: uniq(v.map(r => r[1])),
    puntos: uniq(v.map(r => r[2])),
    materiales: uniq(v.map(r => r[3]))
  });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);

    if (sh.getLastRow() === 0) {
      sh.appendRow(["Fecha de registro","Nombre","Cargo","Fecha","Punto de venta","Tipo de material","Cantidad","Foto"]);
      sh.getRange(1,1,1,8).setFontWeight("bold").setBackground("#1f3a5f").setFontColor("#ffffff");
      sh.setFrozenRows(1);
    }

    let urlFoto = "";
    if (d.foto) {
      const it = DriveApp.getFoldersByName(CARPETA_FOTOS);
      const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FOTOS);
      const bytes = Utilities.base64Decode(d.foto.split(",")[1]);
      const nombreArchivo = (d.punto + "_" + d.fecha + "_" + Date.now() + ".jpg").replace(/[^\w.\-]/g, "_");
      const archivo = carpeta.createFile(Utilities.newBlob(bytes, "image/jpeg", nombreArchivo));
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFoto = archivo.getUrl();
    }

    sh.appendRow([new Date(), d.nombre, d.cargo, d.fecha, d.punto, d.material, Number(d.cantidad), urlFoto]);
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  }
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
