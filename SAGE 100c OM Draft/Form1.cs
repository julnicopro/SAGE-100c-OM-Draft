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
            // Début traitement         
            if (gc.IsOpen)
            {
                MessageBox.Show("Conenxion OK.");
            }

            gc.Close();
        }
    }
}