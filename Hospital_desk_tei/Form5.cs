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
    public partial class Form5 : Form
    {
        List<string> d = new List<string>();
        bool _Rc;
        public bool Rc { get { return _Rc; } }

        public Form5()
        {

            InitializeComponent();
        }
        public Form5(List<string> g)
        {
            d = new List<string>(g);
            InitializeComponent();
            _Rc = false;
        }

        public string FIO
        {
            get { return comboBox1.Text.ToString(); }
            set { comboBox1.Text = value; }

        }
        public string Pr
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
        public DateTime Date
        {
            get { return Convert.ToDateTime(dateTimePicker1.Text); }
            set { dateTimePicker1.Text = value.ToString(); }
        }

        public int Cost
        {
            get { return Convert.ToInt32(textBox2.Text); }
            set { textBox2.Text = value.ToString(); }
        }
        private void Form5_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < d.Count; i++)
                comboBox1.Items.Add(d[i]);

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
