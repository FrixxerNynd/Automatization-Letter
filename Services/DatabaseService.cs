using System;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;
using GeneradorDocumentosSQL.Models;

namespace GeneradorDocumentosSQL.Services
{
    public class DatabaseService
    {
        private readonly string _connectionString;
        private readonly ILogger<DatabaseService> _logger;

        public DatabaseService(string connectionString, ILogger<DatabaseService> logger)
        {
            _connectionString = connectionString;
            _logger = logger;
        }

        public async Task<Cliente?> ObtenerDatosParaCartaAsync(int ultimoId)
        {
            try
            {
                await using var connection = new SqlConnection(_connectionString);
                await connection.OpenAsync();

                string query = @"
                    SELECT TOP 1
                        c.idcliente,
                        c.nombre,
                        t.fecha,   -- reemplaza con tus columnas 
                        t.folio
                    FROM  [tempcheques] t
                    INNER JOIN [clientes] c ON t.idcliente = c.idcliente
                    WHERE t.folio > @ultimoId
                    ORDER BY t.folio ASC";

                await using var command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@ultimoId", ultimoId);

                await using var reader = await command.ExecuteReaderAsync();

                if (await reader.ReadAsync())
                {
                    // Lectura segura del idticket
                    if (!int.TryParse(reader["folio"]?.ToString(), out int Folio))
                    {
                        _logger.LogWarning("El campo idticket no pudo convertirse a entero. Valor: {Valor}", reader["idticket"]);
                        return null;
                    }

                    var cliente = new Cliente
                    {
                        Folio = Folio,
                        IdCliente = reader["idcliente"] == DBNull.Value
                            ? string.Empty
                            : reader["idcliente"].ToString()!,

                        // Lectura segura de strings (por si algún campo es NULL en BD)
                        Nombre = reader["nombre"] == DBNull.Value
                            ? string.Empty
                            : reader["nombre"].ToString()!,

                        // Agrega aquí el resto de tus campos con el mismo patrón
                        FechaAlta = reader["fecha"] == DBNull.Value
                            ? DateTime.Now
                            : Convert.ToDateTime(reader["fecha"]),
                    };

                    return cliente;
                }

                return null;
            }
            catch (SqlException ex)
            {
                _logger.LogError(ex, "Error de SQL al obtener datos para carta. UltimoId: {UltimoId}", ultimoId);
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error inesperado en ObtenerDatosParaCartaAsync. UltimoId: {UltimoId}", ultimoId);
                throw;
            }
        }
    }
}