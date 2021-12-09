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
    public partial class Form2 : Form
    {
        bool _Rc;
        public bool Rc { get { return _Rc; } }

        public string nm
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
        public Form2()
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
