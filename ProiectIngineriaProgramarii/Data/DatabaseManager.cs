using System;
using System.IO;
using Microsoft.Data.Sqlite;
using ProiectIngineriaProgramarii.Models;

namespace ProiectIngineriaProgramarii.Data
{
    public class DatabaseManager
    {
        private readonly string _connectionString;

        public DatabaseManager()
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ingineria.db");
            _connectionString = $"Data Source={dbPath}";
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS Produse (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Nume TEXT NOT NULL,
                    Descriere TEXT,
                    Pret REAL NOT NULL,
                    StocDisponibil INTEGER NOT NULL,
                    DataAdaugare TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS Clienti (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Nume TEXT NOT NULL,
                    Prenume TEXT NOT NULL,
                    Email TEXT,
                    Telefon TEXT,
                    Adresa TEXT,
                    DataInregistrare TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS Facturi (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    NumarFactura TEXT NOT NULL UNIQUE,
                    DataEmitere TEXT NOT NULL,
                    ClientId INTEGER NOT NULL,
                    Subtotal REAL NOT NULL,
                    TVA REAL NOT NULL,
                    Total REAL NOT NULL,
                    Observatii TEXT,
                    Status TEXT NOT NULL,
                    FOREIGN KEY (ClientId) REFERENCES Clienti(Id)
                );

                CREATE TABLE IF NOT EXISTS ItemuriFactura (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FacturaId INTEGER NOT NULL,
                    ProdusId INTEGER NOT NULL,
                    NumeProdus TEXT NOT NULL,
                    Cantitate INTEGER NOT NULL,
                    PretUnitar REAL NOT NULL,
                    Subtotal REAL NOT NULL,
                    FOREIGN KEY (FacturaId) REFERENCES Facturi(Id),
                    FOREIGN KEY (ProdusId) REFERENCES Produse(Id)
                );
            ";
            command.ExecuteNonQuery();
        }

        public SqliteConnection GetConnection()
        {
            return new SqliteConnection(_connectionString);
        }
    }
}
