using System;
using System.Collections.Generic;
using ProiectIngineriaProgramarii.Models;

namespace ProiectIngineriaProgramarii.Data
{
    public class DataSeeder
    {
        private readonly DatabaseManager _dbManager;
        private readonly ClientRepository _clientRepository;
        private readonly ProdusRepository _produsRepository;
        private readonly FacturaRepository _facturaRepository;

        public DataSeeder(DatabaseManager dbManager)
        {
            _dbManager = dbManager;
            _clientRepository = new ClientRepository(_dbManager);
            _produsRepository = new ProdusRepository(_dbManager);
            _facturaRepository = new FacturaRepository(_dbManager, _clientRepository);
        }

        public void SeedData()
        {
            // Verifica daca exista deja date
            if (_clientRepository.GetAll().Count > 0 || _produsRepository.GetAll().Count > 0)
            {
                return; // Datele exista deja, nu mai adaugam
            }

            // Adauga 10 clienti
            SeedClienti();

            // Adauga 20 produse
            SeedProduse();

            // Adauga cateva facturi pentru testare
            SeedFacturi();
        }

        private void SeedClienti()
        {
            var clienti = new List<Client>
            {
                new Client
                {
                    Nume = "Popescu",
                    Prenume = "Ion",
                    Email = "ion.popescu@email.com",
                    Telefon = "0722123456",
                    Adresa = "Str. Victoriei, nr. 10, Bucuresti",
                    DataInregistrare = DateTime.Now.AddMonths(-6)
                },
                new Client
                {
                    Nume = "Ionescu",
                    Prenume = "Maria",
                    Email = "maria.ionescu@email.com",
                    Telefon = "0733234567",
                    Adresa = "Bd. Unirii, nr. 25, Bucuresti",
                    DataInregistrare = DateTime.Now.AddMonths(-5)
                },
                new Client
                {
                    Nume = "Georgescu",
                    Prenume = "Andrei",
                    Email = "andrei.georgescu@email.com",
                    Telefon = "0744345678",
                    Adresa = "Str. Libertatii, nr. 5, Cluj-Napoca",
                    DataInregistrare = DateTime.Now.AddMonths(-4)
                },
                new Client
                {
                    Nume = "Constantinescu",
                    Prenume = "Elena",
                    Email = "elena.const@email.com",
                    Telefon = "0755456789",
                    Adresa = "Calea Mosilor, nr. 123, Bucuresti",
                    DataInregistrare = DateTime.Now.AddMonths(-4)
                },
                new Client
                {
                    Nume = "Dumitrescu",
                    Prenume = "Mihai",
                    Email = "mihai.dumitrescu@email.com",
                    Telefon = "0766567890",
                    Adresa = "Str. Pacii, nr. 45, Timisoara",
                    DataInregistrare = DateTime.Now.AddMonths(-3)
                },
                new Client
                {
                    Nume = "Vasilescu",
                    Prenume = "Ana",
                    Email = "ana.vasilescu@email.com",
                    Telefon = "0777678901",
                    Adresa = "Bd. Revolutiei, nr. 78, Iasi",
                    DataInregistrare = DateTime.Now.AddMonths(-3)
                },
                new Client
                {
                    Nume = "Marinescu",
                    Prenume = "Gheorghe",
                    Email = "gheorghe.marinescu@email.com",
                    Telefon = "0788789012",
                    Adresa = "Str. Florilor, nr. 12, Constanta",
                    DataInregistrare = DateTime.Now.AddMonths(-2)
                },
                new Client
                {
                    Nume = "Stanescu",
                    Prenume = "Ioana",
                    Email = "ioana.stanescu@email.com",
                    Telefon = "0799890123",
                    Adresa = "Calea Dorobantilor, nr. 56, Bucuresti",
                    DataInregistrare = DateTime.Now.AddMonths(-2)
                },
                new Client
                {
                    Nume = "Radulescu",
                    Prenume = "Vlad",
                    Email = "vlad.radulescu@email.com",
                    Telefon = "0721901234",
                    Adresa = "Str. Primaverii, nr. 89, Brasov",
                    DataInregistrare = DateTime.Now.AddMonths(-1)
                },
                new Client
                {
                    Nume = "Niculescu",
                    Prenume = "Cristina",
                    Email = "cristina.niculescu@email.com",
                    Telefon = "0732012345",
                    Adresa = "Bd. Independentei, nr. 34, Ploiesti",
                    DataInregistrare = DateTime.Now.AddMonths(-1)
                }
            };

            foreach (var client in clienti)
            {
                _clientRepository.Add(client);
            }
        }

        private void SeedProduse()
        {
            var produse = new List<Produs>
            {
                new Produs
                {
                    Nume = "Laptop Dell Inspiron 15",
                    Descriere = "Laptop performant pentru birou si gaming",
                    Pret = 3499.99m,
                    StocDisponibil = 25,
                    DataAdaugare = DateTime.Now.AddMonths(-3)
                },
                new Produs
                {
                    Nume = "Monitor Samsung 27\"",
                    Descriere = "Monitor LED Full HD",
                    Pret = 899.99m,
                    StocDisponibil = 40,
                    DataAdaugare = DateTime.Now.AddMonths(-3)
                },
                new Produs
                {
                    Nume = "Tastatura Mecanica RGB",
                    Descriere = "Tastatura gaming cu iluminare RGB",
                    Pret = 349.99m,
                    StocDisponibil = 60,
                    DataAdaugare = DateTime.Now.AddMonths(-2)
                },
                new Produs
                {
                    Nume = "Mouse Wireless Logitech",
                    Descriere = "Mouse ergonomic wireless",
                    Pret = 149.99m,
                    StocDisponibil = 80,
                    DataAdaugare = DateTime.Now.AddMonths(-2)
                },
                new Produs
                {
                    Nume = "Casti Audio Sony WH-1000XM4",
                    Descriere = "Casti wireless cu noise cancelling",
                    Pret = 1299.99m,
                    StocDisponibil = 30,
                    DataAdaugare = DateTime.Now.AddMonths(-2)
                },
                new Produs
                {
                    Nume = "Imprimanta HP LaserJet",
                    Descriere = "Imprimanta laser monocrom",
                    Pret = 899.99m,
                    StocDisponibil = 20,
                    DataAdaugare = DateTime.Now.AddMonths(-2)
                },
                new Produs
                {
                    Nume = "Router TP-Link Gigabit",
                    Descriere = "Router wireless dual-band",
                    Pret = 299.99m,
                    StocDisponibil = 45,
                    DataAdaugare = DateTime.Now.AddMonths(-1)
                },
                new Produs
                {
                    Nume = "Webcam Logitech Full HD",
                    Descriere = "Webcam 1080p pentru conferinte",
                    Pret = 449.99m,
                    StocDisponibil = 35,
                    DataAdaugare = DateTime.Now.AddMonths(-1)
                },
                new Produs
                {
                    Nume = "SSD Samsung 1TB",
                    Descriere = "SSD NVMe M.2 ultra-rapid",
                    Pret = 549.99m,
                    StocDisponibil = 50,
                    DataAdaugare = DateTime.Now.AddMonths(-1)
                },
                new Produs
                {
                    Nume = "HDD External Seagate 2TB",
                    Descriere = "Hard disk extern portabil",
                    Pret = 349.99m,
                    StocDisponibil = 55,
                    DataAdaugare = DateTime.Now.AddMonths(-1)
                },
                new Produs
                {
                    Nume = "Procesor Intel Core i7",
                    Descriere = "Procesor desktop gen 12",
                    Pret = 1899.99m,
                    StocDisponibil = 15,
                    DataAdaugare = DateTime.Now.AddDays(-20)
                },
                new Produs
                {
                    Nume = "Placa Video NVIDIA RTX 3060",
                    Descriere = "Placa video gaming 12GB",
                    Pret = 2499.99m,
                    StocDisponibil = 12,
                    DataAdaugare = DateTime.Now.AddDays(-20)
                },
                new Produs
                {
                    Nume = "Memorie RAM DDR4 16GB",
                    Descriere = "Kit memorie RAM 2x8GB 3200MHz",
                    Pret = 299.99m,
                    StocDisponibil = 70,
                    DataAdaugare = DateTime.Now.AddDays(-15)
                },
                new Produs
                {
                    Nume = "Placa de Baza ASUS",
                    Descriere = "Placa de baza ATX socket 1700",
                    Pret = 799.99m,
                    StocDisponibil = 25,
                    DataAdaugare = DateTime.Now.AddDays(-15)
                },
                new Produs
                {
                    Nume = "Sursa Modulara 750W",
                    Descriere = "Sursa PC 80+ Gold certificata",
                    Pret = 549.99m,
                    StocDisponibil = 30,
                    DataAdaugare = DateTime.Now.AddDays(-10)
                },
                new Produs
                {
                    Nume = "Carcasa PC RGB",
                    Descriere = "Carcasa gaming cu geam lateral",
                    Pret = 399.99m,
                    StocDisponibil = 28,
                    DataAdaugare = DateTime.Now.AddDays(-10)
                },
                new Produs
                {
                    Nume = "Cooler CPU Noctua",
                    Descriere = "Cooler performant silentios",
                    Pret = 349.99m,
                    StocDisponibil = 40,
                    DataAdaugare = DateTime.Now.AddDays(-5)
                },
                new Produs
                {
                    Nume = "Microfon USB Blue Yeti",
                    Descriere = "Microfon condensator pentru streaming",
                    Pret = 649.99m,
                    StocDisponibil = 22,
                    DataAdaugare = DateTime.Now.AddDays(-5)
                },
                new Produs
                {
                    Nume = "Boxe 2.1 Creative",
                    Descriere = "Sistem audio cu subwoofer",
                    Pret = 449.99m,
                    StocDisponibil = 33,
                    DataAdaugare = DateTime.Now.AddDays(-3)
                },
                new Produs
                {
                    Nume = "Hub USB-C 7-in-1",
                    Descriere = "Adaptor multifunctional USB-C",
                    Pret = 199.99m,
                    StocDisponibil = 65,
                    DataAdaugare = DateTime.Now.AddDays(-1)
                }
            };

            foreach (var produs in produse)
            {
                _produsRepository.Add(produs);
            }
        }

        private void SeedFacturi()
        {
            // Obtine clientii si produsele adaugate
            var clienti = _clientRepository.GetAll();
            var produse = _produsRepository.GetAll();

            if (clienti.Count == 0 || produse.Count == 0)
                return;

            var random = new Random();

            // Genereaza 15 facturi pentru testare
            for (int i = 1; i <= 15; i++)
            {
                var factura = new Factura
                {
                    NumarFactura = $"FAC-{DateTime.Now.Year}-{i:D4}",
                    DataEmitere = DateTime.Now.AddMonths(-random.Next(1, 6)).AddDays(-random.Next(0, 30)),
                    ClientId = clienti[random.Next(clienti.Count)].Id,
                    Status = "Emisa"
                };

                // Adauga 2-5 produse random la fiecare factura
                int numarProduse = random.Next(2, 6);
                for (int j = 0; j < numarProduse; j++)
                {
                    var produs = produse[random.Next(produse.Count)];
                    var item = new ItemFactura
                    {
                        ProdusId = produs.Id,
                        NumeProdus = produs.Nume,
                        Cantitate = random.Next(1, 4),
                        PretUnitar = produs.Pret
                    };
                    item.CalculeazaSubtotal();
                    factura.Itemi.Add(item);
                }

                factura.CalculeazaTotaluri();
                _facturaRepository.Add(factura);
            }
        }
    }
}
