using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Hospital_desk_tei
{
    public partial class Form3 : Form
    {
        bool _Rc;
        public bool Rc { get { return _Rc; } }

        public int Nm
        {
            get { return Convert.ToInt32(textBox1.Text); }
            set { textBox1.Text = Convert.ToString(value); }
        }
        public int Vm
        {
            get { return Convert.ToInt32(domainUpDown1.Text); }
            set { domainUpDown1.Text = Convert.ToString(value); }
        }
        public Form3()
        {
            InitializeComponent();
            _Rc = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _Rc = true;
            Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            _Rc = false;
            Close();

        }
    }
}
