using OfficeOpenXml;
using System.Data;
using System.Data.SqlClient;

namespace app
{
    public class DAL
    {


        public void ChargerDonnees(string NomFichier, string CodeFamille, string NomPole, int Periode, int Annee)
        {
            SqlConnection connection = null;
            try
            {
                // Vérification des paramètres
                if (string.IsNullOrEmpty(NomFichier) || string.IsNullOrEmpty(CodeFamille) || string.IsNullOrEmpty(NomPole))
                    throw new ArgumentException("Les paramètres ne peuvent pas être null ou vides.");

                // Vérification de l'existence du fichier
                if (!File.Exists(NomFichier))
                    throw new FileNotFoundException("Le fichier spécifié n'existe pas.");

                // Connexion à la base de donnéestt
                connection = new SqlConnection("Data Source=srvmonitoring;Initial Catalog=Monitoring;User ID=Monitoring;Password=Monitoring@23;");
                connection.Open();
                {

                    using (ExcelPackage package = new ExcelPackage(new FileInfo(NomFichier)))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        // Récupération de la feuille de calcul
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();


                        if (worksheet != null)
                        {
                            // Lecture des données
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {

                                // Récupération des valeurs de la ligne
                                string numeroDossierTps = worksheet.Cells[row, 1].Value?.ToString();
                                string nomOperateur = worksheet.Cells[row, 2].Value?.ToString();
                                string nomBeneficiaire = worksheet.Cells[row, 3].Value?.ToString();
                                DateTime dateDossierTps;
                                if (!DateTime.TryParse(worksheet.Cells[row, 4].Value?.ToString(), out dateDossierTps) || dateDossierTps == DateTime.MinValue)
                                    throw new ArgumentException($"La date de la ligne {row} est invalide.");
                                string codeFormulaire = worksheet.Cells[row, 5].Value?.ToString();
                                string niveauExecution = worksheet.Cells[row, 6].Value?.ToString();
                                DateTime dateRequete;
                                if (!DateTime.TryParse(worksheet.Cells[row, 7].Value?.ToString(), out dateRequete) || dateRequete == DateTime.MinValue)
                                    throw new ArgumentException($"La date de la ligne {row} est invalide.");
                                DateTime dateRetour;
                                if (!DateTime.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out dateRetour) || dateRetour == DateTime.MinValue)
                                    throw new ArgumentException($"La date de la ligne {row} est invalide.");
                                string signataire = worksheet.Cells[row, 9].Value?.ToString();
                                string importationOuExportation = worksheet.Cells[row, 10].Value?.ToString();

                                // Vérification des valeurs lues
                                if (string.IsNullOrEmpty(numeroDossierTps) || string.IsNullOrEmpty(nomOperateur) || string.IsNullOrEmpty(nomBeneficiaire)
                                    || string.IsNullOrEmpty(codeFormulaire) || string.IsNullOrEmpty(niveauExecution) || string.IsNullOrEmpty(signataire)
                                    || string.IsNullOrEmpty(importationOuExportation))
                                    throw new ArgumentException($"Une ou plusieurs valeurs de la ligne {row} sont null ou vides.");

                                // Insertion dans la base de données
                                string query = "INSERT INTO UNEPARTIEDEJOINDRE (NumeroDossierTps, CodeFormulaire, NiveauExecution, DateRequete, DateRetour, NomPole, Periode, Annee, CodeFamille, Signataire, ImportationOuExportation) " +
                                                "VALUES (@NumeroDossierTps, @CodeFormulaire, @NiveauExecution, @DateRequete, @DateRetour, @NomPole, @Periode, @Annee, @CodeFamille, @Signataire, @ImportationOuExportation)";
                                using (SqlCommand commande = new SqlCommand(query, connection))
                                {
                                    // Verifier si les parametres ne sont pas null ou vide avant l'insertion dans la base de donnée 
                                    if (string.IsNullOrEmpty(numeroDossierTps))
                                    {
                                        throw new Exception("Le numéro de dossier TPS ne peut pas être null ou vide.");
                                    }
                                    if (string.IsNullOrEmpty(nomOperateur))
                                    {
                                        throw new Exception("Le nom de l'opérateur ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(nomBeneficiaire))
                                    {
                                        throw new Exception("Le nom du bénéficiaire ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(codeFormulaire))
                                    {
                                        throw new Exception("Le code formulaire ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(niveauExecution))
                                    {
                                        throw new Exception("Le niveau d'exécution ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(signataire))
                                    {
                                        throw new Exception("Le signataire ne peut pas être null ou vide.");
                                    }

                                    if (string.IsNullOrEmpty(importationOuExportation))
                                    {
                                        throw new Exception("La valeur d'importation ou d'exportation ne peut pas être null ou vide.");
                                    }

                                    if (dateDossierTps == default(DateTime))
                                    {
                                        throw new Exception("La date de dossier TPS n'est pas valide.");
                                    }

                                    if (dateRequete == default(DateTime))
                                    {
                                        throw new Exception("La date de requête n'est pas valide.");
                                    }

                                    if (dateRetour == default(DateTime))
                                    {
                                        throw new Exception("La date de retour n'est pas valide.");
                                    }
                                    // Ajouter les paramètres
                                    commande.Parameters.AddWithValue("@NumeroDossierTps", numeroDossierTps);
                                    commande.Parameters.AddWithValue("@CodeFormulaire", codeFormulaire);
                                    commande.Parameters.AddWithValue("@NiveauExecution", niveauExecution);
                                    commande.Parameters.AddWithValue("@DateRequete", dateRequete);
                                    commande.Parameters.AddWithValue("@DateRetour", dateRetour);
                                    commande.Parameters.AddWithValue("@NomPole", NomPole);
                                    commande.Parameters.AddWithValue("@Periode", Periode);
                                    commande.Parameters.AddWithValue("@Annee", Annee);
                                    commande.Parameters.AddWithValue("@CodeFamille", CodeFamille);
                                    commande.Parameters.AddWithValue("@Signataire", signataire);
                                    commande.Parameters.AddWithValue("@ImportationOuExportation", importationOuExportation);

                                    // Exécution de la commande SQL
                                    int rowsAffected = commande.ExecuteNonQuery();

                                    // Vérifier si l'insertion a réussi
                                    if (rowsAffected > 0)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        throw new Exception("L'insertion des données dans la base de données a échoué.");
                                    }

                                }

                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                // Gestion de l'exception
                throw new Exception("Une erreur est survenue lors du chargement des données : " + ex.Message);
            }
            finally
            {
                // Fermeture de la connexion à la base de données
                if (connection != null && connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }
    }
}