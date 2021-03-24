using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoApp
{
    public partial class Flash : Form
    {
        public Flash()
        {
            InitializeComponent();
        }

        private void loader_timer_Tick(object sender, EventArgs e)
        {
            panel2.Width += 3;
            if (panel2.Width>=800) {
                loader_timer.Stop();
                Home home_scrn = new Home();
                home_scrn.Show();
                this.Hide();
            }
        }

   
    }
}
