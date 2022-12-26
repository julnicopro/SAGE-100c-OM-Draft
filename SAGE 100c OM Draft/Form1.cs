using Objets100cLib;

namespace SAGE_100c_OM_Draft
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Debut
            IBSCIALApplication3 gc = new BSCIALApplication100c();
            IBSCPTAApplication3 compta = new BSCPTAApplication100c();

            compta.Name = @"C:\Bases\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\Bases\bijou.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"";

            gc.CptaApplication = (BSCPTAApplication100c)compta;
            gc.Open();
            // D�but traitement         
            if (gc.IsOpen)
            {
                MessageBox.Show("Connexion OK.");
            }

            gc.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Debut
            IBSCIALApplication3 gc = new BSCIALApplication100c();
            IBSCPTAApplication3 compta = new BSCPTAApplication100c();

            compta.Name = @"C:\Bases\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\Bases\bijou.gcm";
            gc.Loggable.UserName = @"<administrateur>";
            gc.Loggable.UserPwd = @"";

            gc.CptaApplication = (BSCPTAApplication100c)compta;
            gc.Open();
            // D�but traitement         
            if (gc.IsOpen)
            {
                //MessageBox.Show("Connexion OK.");
                IBOJournal3 monJournal = compta.FactoryJournal.ReadNumero("VTE");
                DateTime dateJour = DateTime.Today;
                switch (monJournal.TypePeriodCloture[dateJour])
                {
                    case PeriodClotureType.PeriodClotureTotal:
                        MessageBox.Show("Journal clotur� totalement � la date du " + dateJour + ".");
                        break;
                    case PeriodClotureType.PeriodCloturePartial:
                        MessageBox.Show("Journal clotur� partiellement � la date du " + dateJour + ".");
                        break;
                    case PeriodClotureType.PeriodClotureNone:
                        MessageBox.Show("Journal non clotur� � la date du " + dateJour + ".");
                        break;
                    default:
                        MessageBox.Show("Etat cl�ture inconnu � la date du " + dateJour + ".");
                        break;
                }
            }

            gc.Close();
        }
    }
}