using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Objets100cLib;

namespace Référence_Objets_Métier_SAGE
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // Debut
            IBSCIALApplication3 gc = new BSCIALApplication100c();
            IBSCPTAApplication3 compta = new BSCPTAApplication100c();

            compta.Name = @"C:\base\bijou.mae";
            compta.Loggable.UserName = @"<administrateur>";
            compta.Loggable.UserPwd = @"";

            gc.Name = @"C:\base\bijou.gcm";
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