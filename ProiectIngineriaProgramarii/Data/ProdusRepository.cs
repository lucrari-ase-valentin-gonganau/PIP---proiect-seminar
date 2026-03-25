using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using ProiectIngineriaProgramarii.Models;

namespace ProiectIngineriaProgramarii.Data
{
    public class ProdusRepository
    {
        private readonly DatabaseManager _dbManager;

        public ProdusRepository(DatabaseManager dbManager)
        {
            _dbManager = dbManager;
        }

        public List<Produs> GetAll()
        {
            var produse = new List<Produs>();

            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Nume, Descriere, Pret, StocDisponibil, DataAdaugare FROM Produse";

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                produse.Add(new Produs
                {
                    Id = reader.GetInt32(0),
                    Nume = reader.GetString(1),
                    Descriere = reader.IsDBNull(2) ? "" : reader.GetString(2),
                    Pret = reader.GetDecimal(3),
                    StocDisponibil = reader.GetInt32(4),
                    DataAdaugare = DateTime.Parse(reader.GetString(5))
                });
            }

            return produse;
        }

        public void Add(Produs produs)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                INSERT INTO Produse (Nume, Descriere, Pret, StocDisponibil, DataAdaugare)
                VALUES ($nume, $descriere, $pret, $stoc, $data)
            ";
            command.Parameters.AddWithValue("$nume", produs.Nume);
            command.Parameters.AddWithValue("$descriere", produs.Descriere ?? "");
            command.Parameters.AddWithValue("$pret", produs.Pret);
            command.Parameters.AddWithValue("$stoc", produs.StocDisponibil);
            command.Parameters.AddWithValue("$data", produs.DataAdaugare.ToString("yyyy-MM-dd HH:mm:ss"));

            command.ExecuteNonQuery();
        }

        public void Update(Produs produs)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                UPDATE Produse 
                SET Nume = $nume, Descriere = $descriere, Pret = $pret, 
                    StocDisponibil = $stoc, DataAdaugare = $data
                WHERE Id = $id
            ";
            command.Parameters.AddWithValue("$id", produs.Id);
            command.Parameters.AddWithValue("$nume", produs.Nume);
            command.Parameters.AddWithValue("$descriere", produs.Descriere ?? "");
            command.Parameters.AddWithValue("$pret", produs.Pret);
            command.Parameters.AddWithValue("$stoc", produs.StocDisponibil);
            command.Parameters.AddWithValue("$data", produs.DataAdaugare.ToString("yyyy-MM-dd HH:mm:ss"));

            command.ExecuteNonQuery();
        }

        public void Delete(int id)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = "DELETE FROM Produse WHERE Id = $id";
            command.Parameters.AddWithValue("$id", id);

            command.ExecuteNonQuery();
        }
    }
}
