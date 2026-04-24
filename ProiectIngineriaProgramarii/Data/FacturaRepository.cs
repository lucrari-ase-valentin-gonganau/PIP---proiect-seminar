using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using ProiectIngineriaProgramarii.Models;

namespace ProiectIngineriaProgramarii.Data
{
    public class FacturaRepository
    {
        private readonly DatabaseManager _dbManager;
        private readonly ClientRepository _clientRepository;

        public FacturaRepository(DatabaseManager dbManager, ClientRepository clientRepository)
        {
            _dbManager = dbManager;
            _clientRepository = clientRepository;
        }

        public List<Factura> GetAll()
        {
            var facturi = new List<Factura>();

            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT Id, NumarFactura, DataEmitere, ClientId, Subtotal, TVA, Total, Observatii, Status
                FROM Facturi
                ORDER BY DataEmitere DESC
            ";

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                var factura = new Factura
                {
                    Id = reader.GetInt32(0),
                    NumarFactura = reader.GetString(1),
                    DataEmitere = DateTime.Parse(reader.GetString(2)),
                    ClientId = reader.GetInt32(3),
                    Subtotal = reader.GetDecimal(4),
                    TVA = reader.GetDecimal(5),
                    Total = reader.GetDecimal(6),
                    Observatii = reader.IsDBNull(7) ? "" : reader.GetString(7),
                    Status = reader.GetString(8)
                };

                facturi.Add(factura);
            }

            foreach (var factura in facturi)
            {
                factura.Client = _clientRepository.GetById(factura.ClientId);
                factura.Itemi = GetItemuriFactura(factura.Id);
            }

            return facturi;
        }

        public Factura GetById(int id)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT Id, NumarFactura, DataEmitere, ClientId, Subtotal, TVA, Total, Observatii, Status
                FROM Facturi
                WHERE Id = $id
            ";
            command.Parameters.AddWithValue("$id", id);

            using var reader = command.ExecuteReader();
            if (reader.Read())
            {
                var factura = new Factura
                {
                    Id = reader.GetInt32(0),
                    NumarFactura = reader.GetString(1),
                    DataEmitere = DateTime.Parse(reader.GetString(2)),
                    ClientId = reader.GetInt32(3),
                    Subtotal = reader.GetDecimal(4),
                    TVA = reader.GetDecimal(5),
                    Total = reader.GetDecimal(6),
                    Observatii = reader.IsDBNull(7) ? "" : reader.GetString(7),
                    Status = reader.GetString(8)
                };

                factura.Client = _clientRepository.GetById(factura.ClientId);
                factura.Itemi = GetItemuriFactura(factura.Id);
                return factura;
            }

            return null;
        }

        public void Add(Factura factura)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            using var transaction = connection.BeginTransaction();

            try
            {
                var command = connection.CreateCommand();
                command.CommandText = @"
                    INSERT INTO Facturi (NumarFactura, DataEmitere, ClientId, Subtotal, TVA, Total, Observatii, Status)
                    VALUES ($numar, $data, $clientId, $subtotal, $tva, $total, $obs, $status);
                    SELECT last_insert_rowid();
                ";
                command.Parameters.AddWithValue("$numar", factura.NumarFactura);
                command.Parameters.AddWithValue("$data", factura.DataEmitere.ToString("yyyy-MM-dd HH:mm:ss"));
                command.Parameters.AddWithValue("$clientId", factura.ClientId);
                command.Parameters.AddWithValue("$subtotal", factura.Subtotal);
                command.Parameters.AddWithValue("$tva", factura.TVA);
                command.Parameters.AddWithValue("$total", factura.Total);
                command.Parameters.AddWithValue("$obs", factura.Observatii ?? "");
                command.Parameters.AddWithValue("$status", factura.Status);

                factura.Id = Convert.ToInt32(command.ExecuteScalar());

                foreach (var item in factura.Itemi)
                {
                    AddItemFactura(connection, factura.Id, item);
                }

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        public void Update(Factura factura)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            using var transaction = connection.BeginTransaction();

            try
            {
                var command = connection.CreateCommand();
                command.CommandText = @"
                    UPDATE Facturi 
                    SET NumarFactura = $numar, DataEmitere = $data, ClientId = $clientId,
                        Subtotal = $subtotal, TVA = $tva, Total = $total, 
                        Observatii = $obs, Status = $status
                    WHERE Id = $id
                ";
                command.Parameters.AddWithValue("$id", factura.Id);
                command.Parameters.AddWithValue("$numar", factura.NumarFactura);
                command.Parameters.AddWithValue("$data", factura.DataEmitere.ToString("yyyy-MM-dd HH:mm:ss"));
                command.Parameters.AddWithValue("$clientId", factura.ClientId);
                command.Parameters.AddWithValue("$subtotal", factura.Subtotal);
                command.Parameters.AddWithValue("$tva", factura.TVA);
                command.Parameters.AddWithValue("$total", factura.Total);
                command.Parameters.AddWithValue("$obs", factura.Observatii ?? "");
                command.Parameters.AddWithValue("$status", factura.Status);

                command.ExecuteNonQuery();

                DeleteItemuriFactura(connection, factura.Id);

                foreach (var item in factura.Itemi)
                {
                    AddItemFactura(connection, factura.Id, item);
                }

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        public void Delete(int id)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            using var transaction = connection.BeginTransaction();

            try
            {
                DeleteItemuriFactura(connection, id);

                var command = connection.CreateCommand();
                command.CommandText = "DELETE FROM Facturi WHERE Id = $id";
                command.Parameters.AddWithValue("$id", id);
                command.ExecuteNonQuery();

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        private List<ItemFactura> GetItemuriFactura(int facturaId)
        {
            var itemi = new List<ItemFactura>();

            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT Id, FacturaId, ProdusId, NumeProdus, Cantitate, PretUnitar, Subtotal, UnitateMasura
                FROM ItemuriFactura
                WHERE FacturaId = $facturaId
            ";
            command.Parameters.AddWithValue("$facturaId", facturaId);

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                itemi.Add(new ItemFactura
                {
                    Id = reader.GetInt32(0),
                    FacturaId = reader.GetInt32(1),
                    ProdusId = reader.GetInt32(2),
                    NumeProdus = reader.GetString(3),
                    Cantitate = reader.GetInt32(4),
                    PretUnitar = reader.GetDecimal(5),
                    Subtotal = reader.GetDecimal(6),
                    UnitateMasura = reader.IsDBNull(7) ? "buc" : reader.GetString(7)
                });
            }

            return itemi;
        }

        private void AddItemFactura(SqliteConnection connection, int facturaId, ItemFactura item)
        {
            var command = connection.CreateCommand();
            command.CommandText = @"
                INSERT INTO ItemuriFactura (FacturaId, ProdusId, NumeProdus, Cantitate, PretUnitar, Subtotal, UnitateMasura)
                VALUES ($facturaId, $produsId, $nume, $cant, $pret, $subtotal, $um)
            ";
            command.Parameters.AddWithValue("$facturaId", facturaId);
            command.Parameters.AddWithValue("$produsId", item.ProdusId);
            command.Parameters.AddWithValue("$nume", item.NumeProdus);
            command.Parameters.AddWithValue("$cant", item.Cantitate);
            command.Parameters.AddWithValue("$pret", item.PretUnitar);
            command.Parameters.AddWithValue("$subtotal", item.Subtotal);
            command.Parameters.AddWithValue("$um", item.UnitateMasura ?? "buc");

            command.ExecuteNonQuery();
        }

        private void DeleteItemuriFactura(SqliteConnection connection, int facturaId)
        {
            var command = connection.CreateCommand();
            command.CommandText = "DELETE FROM ItemuriFactura WHERE FacturaId = $facturaId";
            command.Parameters.AddWithValue("$facturaId", facturaId);
            command.ExecuteNonQuery();
        }

        public string GenerareNumarFactura()
        {
            return $"FAC-{DateTime.Now:yyyyMMdd}-{DateTime.Now.Ticks % 10000:D4}";
        }
    }
}
