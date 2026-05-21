const SHEET_PINCH="PINCH_SANITARIO";
const SHEET_USERS="USUARIOS";

function doGet(e){

return jsonResponse_({
ok:true,
message:"API funcionando correctamente"
});

}

function doPost(e){

try{

let payload={};

if(
e &&
e.parameter &&
e.parameter.payload
){

payload=
JSON.parse(
e.parameter.payload
);

}
else if(
e &&
e.postData &&
e.postData.contents
){

payload=
JSON.parse(
e.postData.contents
);

}
else{

return jsonResponse_({
ok:false,
error:"No llegaron datos"
});

}

if(payload.action=="login"){

return jsonResponse_(
loginUser_(
payload.usuario,
payload.password
)
);

}

if(
payload.action==
"guardar_pinch_sanitario"
){

return jsonResponse_(
guardarPinchSanitario_(
payload.rows||[]
)
);

}

return jsonResponse_({
ok:false,
error:"Acción no válida"
});

}catch(error){

return jsonResponse_({
ok:false,
error:error.toString()
});

}

}

function loginUser_(
usuario,
password
){

const ss=
SpreadsheetApp
.getActiveSpreadsheet();

const sh=
ss.getSheetByName(
SHEET_USERS
);

if(!sh){

return{
ok:false,
error:
"No existe hoja USUARIOS"
};

}

const datos=
sh
.getDataRange()
.getValues();

if(
datos.length<2
){

return{
ok:false,
error:
"No hay usuarios"
};

}

const encabezado=
datos.shift();

const idxUsuario=
encabezado.indexOf(
"usuario"
);

const idxPass=
encabezado.indexOf(
"contraseña"
);

const idxSupervisor=
encabezado.indexOf(
"supervisor"
);

const idxRol=
encabezado.indexOf(
"rol"
);

for(
let fila of datos
){

if(

String(
fila[idxUsuario]
).trim()
==
String(
usuario
).trim()

&&

String(
fila[idxPass]
).trim()
==
String(
password
).trim()

){

return{

ok:true,

usuario:
usuario,

supervisor:
fila[
idxSupervisor
]||"",

rol:
fila[
idxRol
]||"supervisor"

};

}

}

return{

ok:false,

error:
"Usuario o contraseña incorrectos"

};

}

function guardarPinchSanitario_(
rows
){

try{

const ss=
SpreadsheetApp
.getActiveSpreadsheet();

let sh=
ss.getSheetByName(
SHEET_PINCH
);

if(!sh){

sh=
ss.insertSheet(
SHEET_PINCH
);

}

if(
rows &&
rows.length
){

sh
.getRange(
sh.getLastRow()+1,
1,
rows.length,
rows[0].length
)
.setValues(rows);

}

return{
ok:true
};

}catch(error){

return{
ok:false,
error:error.toString()
};

}

}

function jsonResponse_(
obj
){

return ContentService
.createTextOutput(
JSON.stringify(obj)
)
.setMimeType(
ContentService
.MimeType.JSON
);

}