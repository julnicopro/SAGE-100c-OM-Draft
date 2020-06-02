using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Objets100Lib;

namespace Référence_Objets_Métier_SAGE
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IBSCIALApplication3 gc = new BSCIALApplication3();
            IBSCPTAApplication3 compta = new BSCPTAApplication3();
           
            compta.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"";

            gc.CptaApplication = (BSCPTAApplication3)compta;
            gc.Open();

            IBODepotFactory2 monDepot = (IBODepotFactory2)gc.FactoryDepot.Create();
            // exécute traitement OM
            //IPMDocTransformer pDocVente = (IPMDocTransformer)gc.Transformation.Vente.CreateProcess_Facturer();
            //pDocVente.AddDocument(gc.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteCommande, "BC00001"));
            //DateTime madate = new DateTime();
            //madate.AddDays(31);
            //madate.AddMonths(12);
            //madate.AddYears(2019);
            //pDocVente.DO_Date = madate;
            //IBOClient3 pClient = (IBOClient3)compta.FactoryClient.Create();
            //pClient.CT_Num = "TEST2";
            //pClient.CT_Intitule = "Mon deuxième client";
            //pClient.SetDefault();
            //pClient.Write();
            IBOEcriture3 MonEcriture;
            //MonEcriture.FactoryEcritureA.
            MonEcriture = (IBOEcriture3)compta.FactoryEcriture.Create();
            #region Date ecriture en datetime
            MonEcriture.Date = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Journal en Objet
            IBOJournal3 JournalEnCours = (IBOJournal3)compta.FactoryJournal.ReadNumero("VTE"); 
            MonEcriture.Journal= JournalEnCours;
            
            //MonEcriture.Journal.JO_Type;
            //MonEcriture.Journal.JO_Contrepartie;
            #endregion
            #region Numéro de pièce en String
            MonEcriture.EC_Piece = MonEcriture.Journal.NextEC_Piece[MonEcriture.Date];
            #endregion
            #region Compte général en objet
            IBOCompteG3 CompteGeneralEncours = (IBOCompteG3)compta.FactoryCompteG.ReadNumero("4110000");
            MonEcriture.CompteG=CompteGeneralEncours;
            #endregion
            #region Compte Auxiliaire en Objet
            IBOTiers3 tiersEnCours = (IBOTiers3)compta.FactoryTiers.ReadNumero("CARAT");
            MonEcriture.Tiers= tiersEnCours;
            #endregion
            #region Libéllé en String
            MonEcriture.EC_Intitule="Mon libellé";
            #endregion
            #region Débit ou Crédit en constante
            MonEcriture.EC_Sens = EcritureSensType.EcritureSensTypeCredit;
            //MonEcriture.EC_Sens = EcritureSensType.EcritureSensTypeDebit;
            #endregion
            #region Montant en Numérique
            //MonEcriture.EC_Montant=99.25;
            MonEcriture.EC_Devise=99.25;
            #endregion
            #region Choix Devise par code ISO
            IBPDevise2 deviseEnCours = (IBPDevise2)compta.FactoryDevise.ReadCodeISO("USD");
            MonEcriture.Devise = deviseEnCours;
            #endregion
            #region Parité Devise
            MonEcriture.EC_Parite = 1.2;
            #endregion
            #region Date échéance
            MonEcriture.EC_Echeance = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Date saisie en datetime
            MonEcriture.DateSaisie = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Règlement Objet Règlement
            IBPReglement3 reglementEnCours = compta.FactoryReglement.ReadIntitule("Chèque");
            MonEcriture.Reglement = reglementEnCours;
            #endregion
            #region Taxe Objet
            IBOTaxe3 taxeEnCours = compta.FactoryTaxe.ReadCode("C20");
            MonEcriture.Taxe= taxeEnCours;
            #endregion
            #region Référence en String
            MonEcriture.EC_Reference = "Ma référence";
            #endregion
            #region Numéro de facture en String
            MonEcriture.EC_RefPiece = "FA00045";
            #endregion
            #region Valeur par défaut sur le reste
            MonEcriture.SetDefault();
            #endregion
            #region Enregistrement
            MonEcriture.Write();
            #endregion
            #region Ecriture2
            IBOEcriture3 MonEcriture2 = (IBOEcriture3)compta.FactoryEcriture.Create();
            #region Numéro de pièce en String
            MonEcriture2.EC_Piece = "P1898";
            #endregion
            #region Journal en Objet
            IBOJournal3 JournalEnCours2 = (IBOJournal3)compta.FactoryJournal.ReadNumero("VTE");
            MonEcriture2.Journal = JournalEnCours2;
            #endregion
            #region Compte général en objet
            IBOCompteG3 CompteGeneralEncours2 = (IBOCompteG3)compta.FactoryCompteG.ReadNumero("701020");
            MonEcriture2.CompteG = CompteGeneralEncours2;
            #endregion
            #region Compte Auxiliaire en Objet
            //IBOTiers3 tiersEnCours2 = (IBOTiers3)compta.FactoryTiers.ReadNumero("CARAT");
            //MonEcriture2.Tiers = tiersEnCours2;
            #endregion
            #region Libéllé en String
            MonEcriture2.EC_Intitule = "Mon libellé";
            #endregion
            #region Débit ou Crédit en constante
            MonEcriture2.EC_Sens = EcritureSensType.EcritureSensTypeDebit;
            //MonEcriture.EC_Sens = EcritureSensType.EcritureSensTypeDebit;
            #endregion
            #region Montant en Numérique
            //MonEcriture2.EC_Montant = 99.25;
            MonEcriture2.EC_Devise = 99.25;
            #endregion
            #region Choix Devise par code ISO
            IBPDevise2 deviseEnCours2 = (IBPDevise2)compta.FactoryDevise.ReadCodeISO("USD");
            MonEcriture2.Devise = deviseEnCours2;
            #endregion
            #region Parité Devise
            MonEcriture2.EC_Parite = 1.2;
            #endregion
            #region Date échéance
            MonEcriture2.EC_Echeance = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Date ecriture en datetime
            MonEcriture2.Date = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Date saisie en datetime
            MonEcriture2.DateSaisie = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Règlement Objet Règlement
            IBPReglement3 reglementEnCours2 = compta.FactoryReglement.ReadIntitule("Chèque");
            MonEcriture2.Reglement = reglementEnCours2;
            #endregion
            #region Taxe Objet
            IBOTaxe3 taxeEnCours2 = compta.FactoryTaxe.ReadCode("C20");
            MonEcriture2.Taxe = taxeEnCours2;
            #endregion
            #region Référence en String
            MonEcriture2.EC_Reference = "Ma référence";
            #endregion
            #region Numéro de facture en String
            MonEcriture2.EC_RefPiece = "FA00045";
            #endregion

            #region Valeur par défaut sur le reste
            MonEcriture2.SetDefault();
            #endregion
            #region Enregistrement
            
            MonEcriture2.WriteDefault();
            #endregion
            #endregion

            #region Ecriture3
            IBOEcriture3 MonEcriture3 = (IBOEcriture3)compta.FactoryEcriture.Create();
            #region Numéro de pièce en String
            MonEcriture3.EC_Piece = "P1899";
            #endregion
            #region Journal en Objet
            IBOJournal3 JournalEnCours3 = (IBOJournal3)compta.FactoryJournal.ReadNumero("BEU");
            MonEcriture3.Journal = JournalEnCours3;
            #endregion
            #region Compte général en objet
            IBOCompteG3 CompteGeneralEncours3 = (IBOCompteG3)compta.FactoryCompteG.ReadNumero("4110000");
            MonEcriture3.CompteG = CompteGeneralEncours3;
            #endregion
            #region Compte Auxiliaire en Objet
            IBOTiers3 tiersEnCours3 = (IBOTiers3)compta.FactoryTiers.ReadNumero("CARAT");
            MonEcriture3.Tiers = tiersEnCours3;
            #endregion
            #region Libéllé en String
            MonEcriture3.EC_Intitule = "Mon libellé";
            #endregion
            #region Débit ou Crédit en constante
            MonEcriture3.EC_Sens = EcritureSensType.EcritureSensTypeDebit;
            //MonEcriture.EC_Sens = EcritureSensType.EcritureSensTypeDebit;
            #endregion
            #region Montant en Numérique
            //MonEcriture3.EC_Montant = 99.25;
            MonEcriture3.EC_Devise = 99.25;
            #endregion
            #region Choix Devise par code ISO
            IBPDevise2 deviseEnCours3 = (IBPDevise2)compta.FactoryDevise.ReadCodeISO("USD");
            MonEcriture3.Devise = deviseEnCours3;
            #endregion
            #region Parité Devise
            MonEcriture3.EC_Parite = 1.2;
            #endregion
            #region Date échéance
            MonEcriture3.EC_Echeance = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Date ecriture en datetime
            MonEcriture3.Date = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Date saisie en datetime
            MonEcriture3.DateSaisie = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Règlement Objet Règlement
            IBPReglement3 reglementEnCours3 = compta.FactoryReglement.ReadIntitule("Chèque");
            MonEcriture3.Reglement = reglementEnCours3;
            #endregion
            #region Taxe Objet
            IBOTaxe3 taxeEnCours3 = compta.FactoryTaxe.ReadCode("C20");
            MonEcriture3.Taxe = taxeEnCours3;
            #endregion
            #region Référence en String
            MonEcriture3.EC_Reference = "Ma référence";
            #endregion
            #region Numéro de facture en String
            MonEcriture3.EC_RefPiece = "FA00045";
            #endregion

            #region Valeur par défaut sur le reste
            MonEcriture3.SetDefault();
            #endregion
            #region Enregistrement
            MonEcriture3.Write();
            #endregion
            #endregion

            #region Ecriture4
            IBOEcriture3 MonEcriture4 = (IBOEcriture3)compta.FactoryEcriture.Create();
            #region Numéro de pièce en String
            MonEcriture4.EC_Piece = "P1899";
            #endregion
            #region Journal en Objet
            IBOJournal3 JournalEnCours4 = (IBOJournal3)compta.FactoryJournal.ReadNumero("BEU");
            MonEcriture4.Journal = JournalEnCours4;
            #endregion
            #region Compte général en objet
            IBOCompteG3 CompteGeneralEncours4 = (IBOCompteG3)compta.FactoryCompteG.ReadNumero("5125");
            MonEcriture4.CompteG = CompteGeneralEncours4;
            #endregion
            #region Compte Auxiliaire en Objet
            //IBOTiers3 tiersEnCours2 = (IBOTiers3)compta.FactoryTiers.ReadNumero("CARAT");
            //MonEcriture2.Tiers = tiersEnCours2;
            #endregion
            #region Libéllé en String
            MonEcriture4.EC_Intitule = "Mon libellé";
            #endregion
            #region Débit ou Crédit en constante
            MonEcriture4.EC_Sens = EcritureSensType.EcritureSensTypeCredit;
            //MonEcriture.EC_Sens = EcritureSensType.EcritureSensTypeDebit;
            #endregion
            #region Montant en Numérique
            //MonEcriture4.EC_Montant = 99.25;
            MonEcriture4.EC_Devise = 99.25;
            #endregion
            #region Choix Devise par code ISO
            IBPDevise2 deviseEnCours4 = (IBPDevise2)compta.FactoryDevise.ReadCodeISO("USD");
            MonEcriture4.Devise = deviseEnCours4;
            #endregion
            #region Parité Devise
            MonEcriture4.EC_Parite = 1.2;
            #endregion
            #region Date échéance
            MonEcriture4.EC_Echeance = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Date ecriture en datetime
            MonEcriture4.Date = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Date saisie en datetime
            MonEcriture4.DateSaisie = Convert.ToDateTime("2015-01-01");
            #endregion
            #region Règlement Objet Règlement
            IBPReglement3 reglementEnCours4 = compta.FactoryReglement.ReadIntitule("Chèque");
            MonEcriture4.Reglement = reglementEnCours4;
            #endregion
            #region Taxe Objet
            IBOTaxe3 taxeEnCours4 = compta.FactoryTaxe.ReadCode("C20");
            MonEcriture4.Taxe = taxeEnCours4;
            #endregion
            #region Référence en String
            MonEcriture4.EC_Reference = "Ma référence";
            #endregion
            #region Numéro de facture en String
            MonEcriture4.EC_RefPiece = "FA00045";
            #endregion

            #region Valeur par défaut sur le reste
            MonEcriture4.SetDefault();
            #endregion
            #region Enregistrement
            //MonEcriture4.Write();
            #endregion
            #endregion


            #region Lettrage
            IPMLettrer pLettrage = (IPMLettrer)compta.CreateProcess_Lettrer();
            pLettrage.AddEcriture(MonEcriture);
            pLettrage.AddEcriture(MonEcriture3);
            pLettrage.type = LettrageType.LettrageTypeLettrageDevise;
            pLettrage.SetLettreDefault();
            if (pLettrage.CanProcess==true) {
                pLettrage.Process();
            }
            else
            {
                MessageBox.Show("Ne peut pas lettrer.");
                IFailInfoCol MesErreurs = pLettrage.Errors;
                foreach(IFailInfo MonErreur in MesErreurs)
                {
                    MessageBox.Show(MonErreur.Text);
                }
            }
            #endregion
            gc.Close();

        }


        private void button3_Click(object sender, EventArgs e)
        {
            // Debut
            IBSCIALApplication3 gc = new BSCIALApplication3();
            IBSCPTAApplication3 compta = new BSCPTAApplication3();

            compta.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"";

            gc.CptaApplication = (BSCPTAApplication3)compta;
            gc.Open();
            // actions
            IBOArticle3 monArticle;
            monArticle = (IBOArticle3)gc.FactoryArticle.Create();
            
            monArticle.AR_Ref = "TESTOM";
            monArticle.AR_Design = "Libellé testOM";
            IBODepot3 monDepot = (IBODepot3)gc.FactoryDepot.ReadIntitule("Code");
            IBOArticleDepot3 maConsultationstock = (IBOArticleDepot3)monArticle.FactoryArticleDepot.ReadDepot(monDepot);
            maConsultationstock = (IBOArticleDepot3)monArticle.FactoryArticleDepot.Create();
            maConsultationstock.Depot = monDepot;
            maConsultationstock.SetDefault();
            maConsultationstock.AS_QteMaxi = 10;
            maConsultationstock.Write();
            // Famille Détail valeur = 0
            IBOFamille3 familleEnCours = gc.FactoryFamille.ReadCode(FamilleType.FamilleTypeDetail, "BIJOUXOR");
            monArticle.Famille = familleEnCours;
            monArticle.SetDefault();
            // Article standard = 0
            monArticle.AR_Type = ArticleType.ArticleTypeStandard;
            // Facbrication = 1, Aucun 0
            monArticle.AR_Nomencl = NomenclatureType.NomenclatureTypeFabrication;
            monArticle.AR_Nomencl = NomenclatureType.NomenclatureTypeFabrication;
            // Lot = 5, CMUP = 2, Aucun = 0
            monArticle.AR_SuiviStock = SuiviStockType.SuiviStockTypeLot;
            monArticle.AR_SuiviStock = SuiviStockType.SuiviStockTypeAucun;
            monArticle.AR_SuiviStock = SuiviStockType.SuiviStockTypeCmup;
            monArticle.AR_SuiviStock = SuiviStockType.SuiviStockTypeFifo;
            monArticle.AR_SuiviStock = SuiviStockType.SuiviStockTypeLifo;
            monArticle.AR_SuiviStock = SuiviStockType.SuiviStockTypeSerie;
            // TTC true HT false
            monArticle.AR_PrixTTC = true;
            monArticle.AR_PrixVen=20.34;
            monArticle.AR_PrixAchat=10.15;
            //Dernier Prix d'achat
            monArticle.AR_PUNet=5;
            //monArticle.AR_Coef=1.2
            IBPUnite uniteEnCours = gc.FactoryUnite.ReadIntitule("Unité");
            monArticle.Unite = uniteEnCours;
            monArticle.AR_Langue1 = "test langue1";
            monArticle.AR_Langue2 = "test langue2";
            monArticle.AR_CodeFiscal = "901849900";
            monArticle.AR_EdiCode = "90184990012345";
            monArticle.AR_Pays = "France";
            IBPProduit2 CatalogueEnCours1;
            IBPProduit2 CatalogueEnCours2;
            IBPProduit2 CatalogueEnCours3;
            IBPProduit2 CatalogueEnCours4;
            CatalogueEnCours1 = gc.FactoryProduit.ReadIntitule("Bijoux");
            CatalogueEnCours2 = CatalogueEnCours1.FactorySousCatalogue.ReadIntitule("Argent");
            CatalogueEnCours3 = CatalogueEnCours2.FactorySousCatalogue.ReadIntitule("Bracelets");
            CatalogueEnCours4 = CatalogueEnCours3.FactorySousCatalogue.ReadIntitule("Bracelets");
 
            monArticle.Catalogue = CatalogueEnCours3;
            monArticle.AR_Sommeil = true;
            monArticle.AR_CodeBarre = "123456";
            monArticle.AR_PoidsBrut = 2;
            monArticle.AR_PoidsNet = 1;
            // Tonne 0, Quintal 1, Kilo 2, Gramme 3, Milligramme 4
            monArticle.AR_UnitePoids = UnitePoidsType.UnitePoidsTypeGramme;
            // Composant 0, Pièce détaché 1, fini 2, semi-fini 3
            monArticle.AR_Nature = FamilleNatureType.FamilleNatureTypeProduitFini;
            monArticle.AR_Nature = FamilleNatureType.FamilleNatureTypeComposant;
            monArticle.AR_Nature = FamilleNatureType.FamilleNatureTypePieceDetachee;
            monArticle.AR_Nature = FamilleNatureType.FamilleNatureTypeProduitSemiFini;
            // Standard 0, Spécifique 1
            monArticle.AR_TypeLancement = LancementType.LancementTypeStandard;
            monArticle.AR_TypeLancement = LancementType.LancementTypeSpecifique;
            // Lancement 0, Maturité 1, Déclin 2
            monArticle.AR_Cycle = CycleType.CycleTypeMaturite;
            monArticle.AR_Cycle = CycleType.CycleTypeLancement;
            monArticle.AR_Cycle = CycleType.CycleTypeDeclin;
            monArticle.AR_NbColis = 2;
            monArticle.AR_Fictif = true;
            // Mineur 0, Majeur 1, Critique 2
            monArticle.AR_Criticite = FamilleCriticiteType.FamilleCriticiteTypeMajeur;
            monArticle.AR_Criticite = FamilleCriticiteType.FamilleCriticiteTypeMineur;
            monArticle.AR_Criticite = FamilleCriticiteType.FamilleCriticiteTypeCritique;
            monArticle.AR_SousTraitance = true;
            IBOArticle3 ArticleSubstitutionEnCours;
            ArticleSubstitutionEnCours = (IBOArticle3)gc.FactoryArticle.ReadReference("LINGOR18");
            monArticle.ArticleSubstitution = ArticleSubstitutionEnCours;
            monArticle.AR_Delai = 3;
            monArticle.AR_DelaiFabrication = 2;
            monArticle.AR_DelaiPeremption = 1;
            monArticle.AR_DelaiSecurite = 4;
            monArticle.AR_Garantie = 360;
            monArticle.AR_Escompte = true;
            monArticle.AR_FactForfait = true;
            monArticle.AR_FactPoids = true;
            monArticle.AR_Contremarque = true;
            monArticle.AR_NotImp = true;
            // Nomenclature QEC
            monArticle.AR_QteOperatoire = 6;
            // Définission de la recette pour x fab
            monArticle.AR_QteComp = 4;
            monArticle.AR_Prevision = true;
            // AR_SaisieVar??????
            monArticle.AR_Photo = "bits jpg";
            monArticle.AR_Publie = true;
            monArticle.AR_HorsStat = true;
            monArticle.AR_CoutStd = 24;
            // Sauvegarde
            monArticle.WriteDefault();
            // Composants
            IBOArticleNomenclature3 Composant = (IBOArticleNomenclature3)monArticle.FactoryArticleNomenclature.List[2];
            //Composant.ArticleComposant.AR_Ref;
            Composant.ArticleComposant = (IBOArticle3)gc.FactoryArticle.ReadReference("COAR001");
            Composant.NO_Qte = 2;
            //Composant.NO_Repartition = 1;
            Composant.NO_Operation = "020";
            Composant.NO_Commentaire = "Commentaire";
            Composant.SousTraitance = true;
            // Fixe 0, Variable 1
            Composant.NO_Type = ComposantType.ComposantTypeVariable;
            Composant.NO_Type = ComposantType.ComposantTypeFixe;
            IBODepot3 depotComposantEnCours = (IBODepot3)gc.FactoryDepot.ReadIntitule("Bijou SA");
            Composant.Depot = depotComposantEnCours;
            

            Composant.WriteDefault();
            // IL par le nom de l'IL
            monArticle.InfoLibre["Marque Commerciale"] = "IL1";
            // Par l'indice de stat
            monArticle.AR_Stat[1] = "Printemps/été";
            monArticle.Write();
            MessageBox.Show("ok");
            // fin

            // Test de dépot
            IBODepot3 monDepot = (IBODepot3)gc.FactoryDepot.ReadIntitule("0007");
            monDepot.DE_Intitule = "test";
            monDepot.Adresse.Adresse = "Adresse";
            monDepot.Adresse.Complement = "Complement";
            monDepot.Adresse.CodePostal = "13300";
            monDepot.Adresse.Ville = "Salon";
            monDepot.Adresse.Pays = "France";
            monDepot.Adresse.CodeRegion = "CTNUM";
            //monDepot.SetDefault();
            monDepot.WriteDefault();
            gc.Close();

            

           

        }

        private void button2_Click(object sender, EventArgs e)
        {
            IBSCIALApplication3 gc = new BSCIALApplication3();
            IBSCPTAApplication3 compta = new BSCPTAApplication3();

            compta.Name = @"C:\Bases\BIOTECHDENTAL.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"YODA";

            gc.Name = @"C:\Bases\BIOTECHDENTAL.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"YODA";

            gc.CptaApplication = (BSCPTAApplication3)compta;
            gc.Open();

            // Document
            IBODocumentAchat3 nDoc = gc.FactoryDocumentAchat.ReadPiece(DocumentType.DocumentTypeAchatCommandeConf, "DOxxxx");
            //IBODocumentAchat3 nDoc = (IBODocumentAchat3)gc.FactoryDocumentAchat.Create();
            //nDoc = mDoc;
            nDoc.Tiers.CT_Num = "CT_Num";
            nDoc.DO_Piece = "CT_Num";
            nDoc.SetDefaultDO_Piece();
            IBOFournisseur3 monFournisseur = compta.FactoryFournisseur.ReadNumero("CT_Num");
            nDoc.SetDefaultFournisseur(monFournisseur);
            nDoc.Write();
            IBODocumentLigne3 ligne = (IBODocumentLigne3)nDoc.FactoryDocumentLigne.Create();
            // test compta
            IBOEcritureA3 MonEcritureA;
            MonEcritureA = (IBOEcritureA3)compta.FactoryEcritureA.Create();

            // Test création d'un document stock transfert interdépot.
            IBODocumentStock3 mDoc;
            IPMDocument pDoc = gc.CreateProcess_Document(DocumentType.DocumentTypeStockVirement);
            IBODepot3 DépotOrigine = gc.FactoryDepot.ReadIntitule("BIOTECH SALON");
            IBODepot3 DépotDestination = gc.FactoryDepot.ReadIntitule("DR DITAC JEAN MARC");
            mDoc = (IBODocumentStock3)pDoc.Document;
            mDoc.SetDefault();
            mDoc.SetDefaultDO_Piece();
            mDoc.DepotOrigine = DépotOrigine;
            mDoc.DepotDestination = DépotDestination;
            mDoc.DO_Date = Convert.ToDateTime("2019-12-31");
            mDoc.DO_Ref = "Prêt DITAC";
            mDoc.WriteDefault();
            // Création du document
            if (pDoc.CanProcess) {
                pDoc.Process();
            }
            // Ajout IL
            mDoc = (IBODocumentStock3)gc.FactoryDocumentStock.ReadPiece(DocumentType.DocumentTypeStockVirement,pDoc.Document.DO_Piece);
            mDoc.InfoLibre["DOCUMENT_ASSOCIE"] = pDoc.Document.DO_Piece;
            mDoc.Write();
            // ajout des lignes
            IBOArticle3 monArticle = gc.FactoryArticle.ReadReference("00TROUSSEKONTACT");
            IPMDocTransferer pLigne = (IPMDocTransferer)gc.CreateProcess_DocTransferer();
            // Lot
            IPMPreleverLot pMonLot = gc.CreateProcess_PreleverLot();
            pMonLot.LigneOrigine = (IBODocumentVenteLigne3)pLigne;
            foreach (IUserLot unLot in pMonLot.UserLots)
            {
                //MessageBox.Show(unLot.Lot.);
            }
            IUserLot pLot = pLigne.UserLotsToUse.AddNew();
            pLot.Lot.Read();
            pLot.Lot = (IBOArticleDepotLot)pLot;
            pLigne.Document = mDoc;
            pLigne.SetDefaultArticle(monArticle, 1);
            pLigne.Process();
            int nb = mDoc.FactoryDocumentLigne.List.Count;
            //foreach(IBODocumentLigne3 MaLigne in MesLignes)
            //{
            //    MaLigne.InfoLibre[37] = mDoc.DO_Piece;
            //    MaLigne.Write();
            //}
            //
            //IBOArticleDepotLot monLot = pLigne.Lot.FactoryArticleDepotLot.ReadNoSerie("26538-61");
            gc.Close();
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Debut
            IBSCIALApplication3 gc = new BSCIALApplication3();
            IBSCPTAApplication3 compta = new BSCPTAApplication3();

            compta.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"";

            gc.CptaApplication = (BSCPTAApplication3)compta;
            gc.Open();

            IBOEcriture3 Ecriture;
            IBOEcritureA3 EcritureA;
            IBOCompteA3 CompteA;
            IBPAnalytique3 Analytique;
            Ecriture = (IBOEcriture3)compta.FactoryEcriture.Create();
            EcritureA =(IBOEcritureA3)Ecriture.FactoryEcritureA.Create();

            Analytique = (IBPAnalytique3)compta.FactoryAnalytique.ReadIntitule("Activité");
            CompteA = (IBOCompteA3)compta.FactoryCompteA.ReadNumero(Analytique,"Section");
            EcritureA.CompteA = CompteA;
            EcritureA.EA_Montant = 16.8;
            EcritureA.SetDefault();
            EcritureA.Write();
            gc.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Debut
            IBSCIALApplication3 gc = new BSCIALApplication3();
            IBSCPTAApplication3 compta = new BSCPTAApplication3();

            compta.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"";

            gc.CptaApplication = (BSCPTAApplication3)compta;
            gc.Open();
            IBODocumentAchat3 monDoc = (IBODocumentAchat3)gc.FactoryDocumentAchat.CreateType(DocumentType.DocumentTypeAchatCommande);
            monDoc.Tiers = compta.FactoryFournisseur.ReadNumero("BILLO");
            monDoc.WriteDefault();
            IBODocumentAchatLigne3 Maligne = (IBODocumentAchatLigne3)gc.FactoryDocumentLigne.Create();
            
            gc.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // Debut
            IBSCIALApplication3 gc = new BSCIALApplication3();
            IBSCPTAApplication3 compta = new BSCPTAApplication3();

            compta.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\Users\Public\Documents\Sage\iSage Entreprise\bijou.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"";

            gc.CptaApplication = (BSCPTAApplication3)compta;
            gc.Open();

            // Choix du client
            IBOClient3 monclient = compta.FactoryClient.ReadNumero("choixClient");
            //nb lignes tarif
            int nb = monclient.FactoryClientTarif.List.Count;
            

            // Choix de l'article
            IBOArticleTarifClient3 monTarifClientEx = (IBOArticleTarifClient3)monclient.FactoryClientTarif.Create();
            IBOArticle3 monArticle = gc.FactoryArticle.ReadReference("choixArticle");
            monTarifClientEx.Article = monArticle;
            // Récupération du contexte
            monTarifClientEx.SetDefault();
            // Choix prix HT ou TTC
            monTarifClientEx.PrixTTC = true;
            // PV
            monTarifClientEx.Prix=25.2;
            // Remise
            monTarifClientEx.Remise.Remise[1].REM_Type = RemiseType.RemiseTypePourcent; // unite = 0, pourcent = 1, unité = 2
            monTarifClientEx.Remise.Remise[1].REM_Valeur = 20;
            // devise
            IBPDevise2 maDevise = compta.FactoryDevise.ReadCodeISO("FR");
            monTarifClientEx.Devise = maDevise;
            // Coefficient
            monTarifClientEx.Coefficient = 0;
            // Calculer le PV selon PA
            monTarifClientEx.Calcul = false;
            monTarifClientEx.Write();
            

            gc.Close();
        }
    }
}