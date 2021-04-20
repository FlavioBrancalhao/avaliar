using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace avaliar
{
    public partial class Form2 : Form
    {
        double milisegundos = 0;
        

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            milisegundos += 15.650;

            if (milisegundos >= 500)
            {
                timer1.Stop();
                this.Close();
            }
        }
    }
}
