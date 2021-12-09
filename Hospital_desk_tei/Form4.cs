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
    public partial class Form4 : Form
    {
        List<string> d = new List<string>();
        bool _Rc;
        public bool Rc { get { return _Rc; } }
        public string Pal
        {
            get { return comboBox1.Text; }
            set { comboBox1.Text = value; }
        }
        public string FIO
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
        public DateTime Date1
        {
            get { return Convert.ToDateTime(dateTimePicker1.Text); }
            set { dateTimePicker1.Text = value.ToString(); }
        }
        public DateTime Date2
        {
            get { return Convert.ToDateTime(dateTimePicker2.Text); }
            set { dateTimePicker2.Text = value.ToString(); }
        }
        public string Di
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }

        public Form4()
        {

            InitializeComponent();
        }
        public Form4(List<string> g)
        {
            d = new List<string>(g);
            InitializeComponent();
            _Rc = false;
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < d.Count; i++)
                comboBox1.Items.Add(d[i]);


        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (Convert.ToDateTime(dateTimePicker2.Text) < Convert.ToDateTime(dateTimePicker1.Text))
            {
                MessageBox.Show("Дата выписки должна быть позже даты поступления!");

            }
            else
            {
                _Rc = true;
                Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _Rc = false;
            Close();
        }
    }
}
