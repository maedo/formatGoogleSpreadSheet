# formatGoogleSpreadSheet
Ayuda a formatear una planilla de Google. Crea las hojas y columnas necesarias para el funcionamiento de un script y establece la ubicación de las columnas por si estas llegan a cambiar dentro del documento. 

Ejemplo de uso:

  fGSS = formatGSS.getFormatGoogleSpreadSheet("CIM", [
    "id", // no cambiar
    "SHARE_TYPE", // no cambiar
    "alternateLink",
    "creationTime",
    "updateTime",
    "name",
    "courseState",
    "description",
    "section",
    "room",
    "ownerId",
  ]);
  fGSS.formatSpreadSheet();
