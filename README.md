# Generador Carta Responsiva

Servicio .NET Worker que monitorea una base de datos SQL Server, detecta nuevos registros de tickets/folios, genera una carta responsiva en formato `.docx` a partir de una plantilla Word y la envía automáticamente a impresión.

---

## Tabla de contenido

- [¿Qué hace este proyecto?](#qué-hace-este-proyecto)
- [Arquitectura técnica](#arquitectura-técnica)
- [Tecnologías y dependencias](#tecnologías-y-dependencias)
- [Requisitos](#requisitos)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Configuración](#configuración)
- [Plantilla Word y placeholders](#plantilla-word-y-placeholders)
- [Instalación y ejecución](#instalación-y-ejecución)
- [Flujo de operación](#flujo-de-operación)
- [Persistencia de estado](#persistencia-de-estado)
- [Logs y monitoreo](#logs-y-monitoreo)
- [Problemas comunes](#problemas-comunes)
- [Seguridad y buenas prácticas](#seguridad-y-buenas-prácticas)
- [Mejoras sugeridas](#mejoras-sugeridas)
- [Licencia](#licencia)

---

## ¿Qué hace este proyecto?

El servicio:

1. Consulta SQL Server para obtener el siguiente registro con `folio > último_folio_procesado`.
2. Genera un `.docx` desde una plantilla (`Templates/carta-responsiva.docx`).
3. Reemplaza placeholders con datos del cliente y fechas.
4. Imprime el documento en una impresora configurada.
5. Elimina el archivo temporal.
6. Guarda el último folio procesado en `estado_vta.json`.
7. Repite el ciclo cada `IntervalMs`.

---

## Arquitectura técnica

Componentes principales:

- **`Program.cs`**
  - Configura DI, logging y servicios.
  - Registra `DatabaseService`, `WordService`, `PrintService` y `Worker`.

- **`Worker.cs`**
  - Orquesta todo el proceso en bucle.
  - Carga/guarda estado (`estado_vta.json`).
  - Limpia archivos `.docx` antiguos según retención.
  - Consulta datos, genera carta, imprime y actualiza estado.

- **`Services/DatabaseService.cs`**
  - Conexión a SQL Server.
  - Ejecuta query para obtener el siguiente ticket/folio.

- **`Services/WordService.cs`**
  - Copia plantilla Word a ruta temporal.
  - Reemplaza placeholders en cuerpo y footers del documento.

- **`Services/PrintService.cs`**
  - Limpia cola de impresión (WMI).
  - Imprime usando `ProcessStartInfo` con `Verb = "printto"`.

- **`Models/Cliente.cs`**
  - Modelo de datos del cliente/ticket para sustituir placeholders.

---

## Tecnologías y dependencias

- .NET Worker Service (`Microsoft.NET.Sdk.Worker`)
- `DocumentFormat.OpenXml`
- `Microsoft.Data.SqlClient`
- `Microsoft.Extensions.Hosting`
- `System.Management`
- `Xceed.Words.NET` (referenciada en proyecto)

> **Nota**: el proyecto está configurado con `TargetFramework: net10.0`.

---

## Requisitos

### Sistema operativo

- **Windows** (impresión con `printto` + WMI `Win32_PrintJob`).

### Software

- .NET SDK compatible con `net10.0`.
- SQL Server accesible desde el host.
- Microsoft Word o asociación de impresión para `.docx` en el equipo (recomendado para `printto`).
- Impresora instalada y visible en Windows con el nombre exacto.

---

## Estructura del proyecto

```text
.
├── Program.cs
├── Worker.cs
├── appsettings.json
├── appsettings.Development.json
├── Generador_Carta_Responsiva.csproj
├── Models/
│   ├── Cliente.cs
│   └── Progreso.cs
├── Services/
│   ├── DatabaseService.cs
│   ├── WordService.cs
│   └── PrintService.cs
└── Templates/
    └── carta-responsiva.docx
```

## Configuración

Configura appsettings.json (o variables de entorno / secretos).

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=...;Database=...;..."
  },
  "WorkerSettings": {
    "TemplatePath": "Templates\\carta-responsiva.docx",
    "PrinterName": "Nombre exacto de la impresora en Windows",
    "IntervalMs": 10000,
    "RetentionDays": 30
  }
}
```

## Campos clave

- **`DefaultConnection`**: conexión a SQL Server.

- **`TemplatePath`**: ruta de plantilla .docx.

- **`OutputPath`**: carpeta para documentos generados (aunque actualmente se usa ruta temporal para impresión).

- **`PrinterName`**: nombre exacto de la impresora.

- **`IntervalMs`**: intervalo de consulta en milisegundos.

- **`RetentionDays`**: días de retención para limpieza de .docx.

## Plantilla Word y placeholders

Placeholders detectados actualmente en el código:

- **`{nombre_cliente}`**

- **`{fecha}`**

- **`{folio}`**

- **`{fecha_impresion}`**

Se reemplazan en:

- Párrafos del cuerpo del documento.

- Párrafos de todos los footers.

> **Recomendación**: evita dividir placeholders en múltiples estilos/runs dentro de Word, para asegurar reemplazo correcto.

## Instalación y ejecución

1. Clonar repositorio

```bash
   git clone <URL_DEL_REPOSITORIO>
   cd Automatization-Letter
```

2. Ajustar configuración
   Edita `appsettings.json`:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=...;Database=...;..." -- Aqui se coloca la cadena de conexion a la base de datos
  },
  "WorkerSettings": {
    "TemplatePath": "Templates\\carta-responsiva.docx", -- Se modifica por la ruta donde se almacena la plantilla
    "PrinterName": "Nombre exacto de la impresora en Windows", -- Se modifica por el nombre de la impresora
    "IntervalMs": 10000, -- Se modifica por el intervalo de tiempo en milisegundos
    "RetentionDays": 30 -- Se modifica por el numero de dias de retencion
  }
}
```

3. Restaurar y compilar

```bash
   dotnet restore
   dotnet build -c Release
```

4. Ejecutar localmente

```bash
   dotnet run
```

## Flujo de operación

1. Al iniciar, el worker carga `estado_vta.json` (si existe).
2. Limpia documentos antiguos en `OutputPath` (extensión `.docx`).
3. Consulta SQL con:

- tablas: `tempcheques` y `clientes`
- condición: ``t.folio > @`ultimoId``
- orden: ascendente, `TOP 1`

4. Si encuentra registro:

- genera archivo temporal `Carta*<folio>*<guid>.docx`
- reemplaza placeholders
- imprime en impresora configurada
- espera 20s
- borra archivo temporal
- actualiza `UltimoId`

5. Espera `IntervalMs` y repite.

## Persistencia de estado

Archivo usado:

- `estado_vta.json` (ubicado junto al binario, `AppContext.BaseDirectory`)

Estructura esperada:

```json
{ "UltimoId": 12345 }
```

Esto evita reprocesar folios ya impresos después de reinicios.

## Logs y monitoreo

El logging está configurado a consola. Eventos comunes:

- inicio del worker
- carga/guardado de estado
- errores SQL o impresión
- generación y eliminación de documentos temporales

Sugerencia para producción:

- redirigir salida a archivo o visor centralizado (Event Viewer, ELK, etc.).

## Problemas comunes

### No imprime

- Verifica `PrinterName` exacto.
- Confirma que `.docx` tiene asociación válida para impresión en el host.
- Revisa si hay permisos/servicio sin sesión interactiva.

### No genera cartas

- Verifica query y columnas en DB:
  - `tempcheques.folio`, `tempcheques.idcliente`, `tempcheques.fecha`
  - `clientes.idcliente`, `clientes.nombre`
- Revisa conexión y cifrado SQL (`Encrypt`, `TrustServerCertificate`).

### Repite o salta folios

- Revisa contenido de `estado_vta.json`.
- Asegura que `folio` sea monotónico e incremental.

## Seguridad y buenas prácticas

- No versionar credenciales reales.
- Mover connection string a secretos/variables de entorno.
- Limitar permisos de usuario del servicio en:
  - SQL Server
  - sistema de archivos
  - impresora
- Respaldar plantilla y validar cambios antes de producción.

## Licencia

Define aquí la licencia del proyecto (por ejemplo, MIT, propietario interno, etc.).
