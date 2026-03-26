using System;

namespace GeneradorDocumentosSQL.Models
{
    public class Cliente
    {
        // Identificador principal
        public int Folio { get; set; }
        public string IdCliente { get; set; }

        // Datos de Identidad
        public string Nombre { get; set; }
        public string PrimerNombre { get; set; }
        public string OtrosNombres { get; set; }
        public string Apellido { get; set; }
        public string SegundoApellido { get; set; }
        public string TratamientoPersonal { get; set; }
        public string RFC { get; set; }
        public string CURP { get; set; }
        public string Contacto { get; set; }

        // Ubicación
        public string Direccion { get; set; }
        public string CodigoPostal { get; set; }
        public string Poblacion { get; set; }
        public string Estado { get; set; }
        public string Pais { get; set; }

        // Comunicación
        public string Email { get; set; }
        public string Telefono1 { get; set; }
        public string Telefono2 { get; set; }
        public string Telefono3 { get; set; }
        public string Telefono4 { get; set; }
        public string Telefono5 { get; set; }

        // Datos Financieros (Usamos decimal para precisión)
        public decimal LimiteDeCredito { get; set; }
        public decimal LimiteCreditoDiario { get; set; }
        public decimal Descuento { get; set; }
        public int DiasVigenciaCredito { get; set; }

        // Fechas
        public DateTime? Cumpleaños { get; set; } // El ? permite nulos
        public DateTime FechaAlta { get; set; }

        // Configuración y SAT
        public string IdPais_SAT { get; set; }
        public string IdRegimen_SAT { get; set; }
        public string FolioFiscal { get; set; }
        public string ObligTributarias { get; set; }
        public int Status { get; set; }

        // Flags/Boleanos (SQL Bit se mapea a bool)
        public bool ProcesadoWeb { get; set; }
        public bool NoCobrarImpuestos { get; set; }
        public bool RetenerImpuesto { get; set; }
        public bool ContemplarPropina { get; set; }

        // Otros IDs y Notas
        public int IdTipoCliente { get; set; }
        public string Giro { get; set; }
        public string Notas { get; set; }
    }
}