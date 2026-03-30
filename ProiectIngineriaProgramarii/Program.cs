using ProiectIngineriaProgramarii.Data;

namespace ProiectIngineriaProgramarii
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            // Initializeaza baza de date si adauga date mockup daca e prima rulare
            var dbManager = new DatabaseManager();
            var seeder = new DataSeeder(dbManager);
            seeder.SeedData();

            Application.Run(new StartForm());
        }
    }
}