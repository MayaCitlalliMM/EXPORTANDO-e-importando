using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EXPORTANDO
{
    public partial class Form1 : Form
    {
        Acciones acc = new Acciones();
        public Form1()
        {
            InitializeComponent();
        }

        private void btnMostrar_Click(object sender, EventArgs e)
        {
            tabladatos.DataSource = null;
            tabladatos.DataSource = acc.MostrarAlumnos();
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            if (acc.ExportarExcel())
            {
                MessageBox.Show("Exportando con exito...");
            }
            else
            {
                MessageBox.Show("Fallo el exportado...");
            }
        }
    }
}
