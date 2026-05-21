/*********************************************************
 SISTEMA ROSAS - GOOGLE APPS SCRIPT
 MODULO CAMPO + TECNICO + PINCH SANITARIO
 CODIGO COMPLETO CORREGIDO
**********************************************************/

const SHEET_CAMPO = "CAMPO";
const SHEET_PINCH = "PINCH_SANITARIO";
const SHEET_USERS = "USUARIOS";

/*********************************************************
 RESPUESTA JSON
**********************************************************/
function jsonResponse_(data){

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

}

/*********************************************************
 DO GET
**********************************************************/
function doGet(e){

  try {

    var action = e.parameter.action || "";

    /*****************************************************
     BASE DE DATOS
    *****************************************************/
    if(action == "base"){

      return jsonResponse_(
        getBaseData_()
      );

    }

    /*****************************************************
     REPORTE DIARIO
    *****************************************************/
    if(action == "reporte_diario"){

      return jsonResponse_(
        getReporteDiario_(e.parameter)
      );

    }

    return jsonResponse_({
      ok:true,
      message:"API funcionando correctamente"
    });

  } catch(err){

    return jsonResponse_({
      ok:false,
      error:String(err)
    });

  }

}

/*********************************************************
 DO POST
**********************************************************/
function doPost(e){

  try {

    // Validar POST
    if(!e || !e.postData){

      return jsonResponse_({
        ok:false,
        message:"No se recibieron datos POST"
      });

    }

    var data = JSON.parse(e.postData.contents);

    if(data.action == "login"){
      return jsonResponse_(loginUser_(data.usuario,data.password));
    }

    /*****************************************************
     GUARDAR PINCH SANITARIO
    *****************************************************/
    if(data.action == "guardar_pinch_sanitario"){

      return jsonResponse_(
        guardarPinchSanitario_(data.rows || [])
      );

    }

    return jsonResponse_({
      ok:false,
      message:"Accion no encontrada"
    });

  } catch(err){

    return jsonResponse_({
      ok:false,
      error:String(err)
    });

  }

}

/*********************************************************
 OBTENER BASE
 COLUMNAS:
 Supervisor | Bloque | Variedad | Modulo
**********************************************************/
function getBaseData_(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getSheetByName(SHEET_CAMPO);

  if(!sh){

    return [];

  }

  var values = sh.getDataRange().getValues();

  if(values.length <= 1){

    return [];

  }

  var headers = values.shift();

  var data = values.map(function(r){

    return {

      Supervisor : r[headers.indexOf("Supervisor")] || "",

      Bloque : r[headers.indexOf("Bloque")] || "",

      Variedad : r[headers.indexOf("Variedad")] || "",

      Modulo : r[headers.indexOf("Modulo")] || ""

    };

  });

  return data;

}

/*********************************************************
 GUARDAR PINCH SANITARIO
**********************************************************/
function guardarPinchSanitario_(rows){

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getSheetByName(SHEET_PINCH);

  /*****************************************************
   CREAR HOJA SI NO EXISTE
  *****************************************************/
  if(!sh){

    sh = ss.insertSheet(SHEET_PINCH);

    sh.appendRow([

      "Fecha",
      "Semana",
      "Supervisor",
      "Modulo",
      "Bloque",
      "Variedad",
      "Lunes",
      "Martes",
      "Miercoles",
      "Jueves",
      "Viernes",
      "Total",
      "Motivo"

    ]);

  }

  /*****************************************************
   GUARDAR FILAS
  *****************************************************/
  rows.forEach(function(r){

    sh.appendRow([

      new Date(),

      r.Semana || "",

      r.Supervisor || "",

      r.Modulo || "",

      r.Bloque || "",

      r.Variedad || "",

      Number(r.Lunes || 0),

      Number(r.Martes || 0),

      Number(r.Miercoles || 0),

      Number(r.Jueves || 0),

      Number(r.Viernes || 0),

      Number(r.Total || 0),

      r.Motivo || ""

    ]);

  });

  return {

    ok:true,
    message:"Datos guardados correctamente"

  };

}

/*********************************************************
 REPORTE DIARIO
 FILTROS:
 fecha
 supervisor
 modulo
**********************************************************/
function getReporteDiario_(params){

  var fechaFiltro = params.fecha || "";

  var supervisorFiltro = params.supervisor || "";

  var moduloFiltro = params.modulo || "";

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getSheetByName(SHEET_PINCH);

  if(!sh){

    return [];

  }

  var values = sh.getDataRange().getValues();

  if(values.length <= 1){

    return [];

  }

  var headers = values.shift();

  var reporte = values.map(function(r){

    return {

      Fecha : Utilities.formatDate(

        new Date(r[headers.indexOf("Fecha")]),

        Session.getScriptTimeZone(),

        "yyyy-MM-dd"

      ),

      Semana : r[headers.indexOf("Semana")] || "",

      Supervisor : r[headers.indexOf("Supervisor")] || "",

      Modulo : r[headers.indexOf("Modulo")] || "",

      Bloque : r[headers.indexOf("Bloque")] || "",

      Variedad : r[headers.indexOf("Variedad")] || "",

      Lunes : Number(
        r[headers.indexOf("Lunes")] || 0
      ),

      Martes : Number(
        r[headers.indexOf("Martes")] || 0
      ),

      Miercoles : Number(
        r[headers.indexOf("Miercoles")] || 0
      ),

      Jueves : Number(
        r[headers.indexOf("Jueves")] || 0
      ),

      Viernes : Number(
        r[headers.indexOf("Viernes")] || 0
      ),

      Total : Number(
        r[headers.indexOf("Total")] || 0
      ),

      Motivo : r[headers.indexOf("Motivo")] || ""

    };

  });

  /*****************************************************
   FILTRAR
  *****************************************************/
  reporte = reporte.filter(function(r){

    var ok = true;

    if(fechaFiltro != ""){

      ok = ok && r.Fecha == fechaFiltro;

    }

    if(supervisorFiltro != ""){

      ok = ok && r.Supervisor == supervisorFiltro;

    }

    if(moduloFiltro != ""){

      ok = ok && r.Modulo == moduloFiltro;

    }

    return ok;

  });

  return reporte;

}

function loginUser_(usuario,password){
 var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
 if(!sh){ return {ok:false,error:"No existe hoja USUARIOS"}; }
 var v=sh.getDataRange().getValues();
 if(v.length<2){ return {ok:false,error:"No hay usuarios"}; }
 var h=v.shift();
 for(var i=0;i<v.length;i++){
  var r=v[i];
  if(String(r[h.indexOf("usuario")])==String(usuario) &&
     String(r[h.indexOf("contraseña")])==String(password)){
      return {ok:true,
      usuario:usuario,
      supervisor:r[h.indexOf("supervisor")]||"",
      rol:r[h.indexOf("rol")]||"supervisor"};
  }
 }
 return {ok:false,error:"Usuario o contraseña incorrectos"};
}
