using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using ProiectIngineriaProgramarii.Models;

namespace ProiectIngineriaProgramarii.Data
{
    public class ClientRepository
    {
        private readonly DatabaseManager _dbManager;

        public ClientRepository(DatabaseManager dbManager)
        {
            _dbManager = dbManager;
        }

        public List<Client> GetAll()
        {
            var clienti = new List<Client>();

            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = "SELECT Id, Nume, Prenume, Email, Telefon, Adresa, DataInregistrare FROM Clienti";

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                clienti.Add(new Client
                {
                    Id = reader.GetInt32(0),
                    Nume = reader.GetString(1),
                    Prenume = reader.GetString(2),
                    Email = reader.IsDBNull(3) ? "" : reader.GetString(3),
                    Telefon = reader.IsDBNull(4) ? "" : reader.GetString(4),
                    Adresa = reader.IsDBNull(5) ? "" : reader.GetString(5),
                    DataInregistrare = DateTime.Parse(reader.GetString(6))
                });
            }

            return clienti;
        }

        public void Add(Client client)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                INSERT INTO Clienti (Nume, Prenume, Email, Telefon, Adresa, DataInregistrare)
                VALUES ($nume, $prenume, $email, $telefon, $adresa, $data)
            ";
            command.Parameters.AddWithValue("$nume", client.Nume);
            command.Parameters.AddWithValue("$prenume", client.Prenume);
            command.Parameters.AddWithValue("$email", client.Email ?? "");
            command.Parameters.AddWithValue("$telefon", client.Telefon ?? "");
            command.Parameters.AddWithValue("$adresa", client.Adresa ?? "");
            command.Parameters.AddWithValue("$data", client.DataInregistrare.ToString("yyyy-MM-dd HH:mm:ss"));

            command.ExecuteNonQuery();
        }

        public void Update(Client client)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                UPDATE Clienti 
                SET Nume = $nume, Prenume = $prenume, Email = $email, 
                    Telefon = $telefon, Adresa = $adresa, DataInregistrare = $data
                WHERE Id = $id
            ";
            command.Parameters.AddWithValue("$id", client.Id);
            command.Parameters.AddWithValue("$nume", client.Nume);
            command.Parameters.AddWithValue("$prenume", client.Prenume);
            command.Parameters.AddWithValue("$email", client.Email ?? "");
            command.Parameters.AddWithValue("$telefon", client.Telefon ?? "");
            command.Parameters.AddWithValue("$adresa", client.Adresa ?? "");
            command.Parameters.AddWithValue("$data", client.DataInregistrare.ToString("yyyy-MM-dd HH:mm:ss"));

            command.ExecuteNonQuery();
        }

        public void Delete(int id)
        {
            using var connection = _dbManager.GetConnection();
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = "DELETE FROM Clienti WHERE Id = $id";
            command.Parameters.AddWithValue("$id", id);

            command.ExecuteNonQuery();
        }
    }
}
