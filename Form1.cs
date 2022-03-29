using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using namespaceJobExcel;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
    

        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)

        {
            DialogVibora();
            FillForm();
        }

        private void FillForm()
        {
            JobExcel JobExcel = new JobExcel();
            JobExcel.FileName = textBox1.Text;
            JobExcel.ReadExcel();

            textBox3.Text = Convert.ToString(JobExcel.buy);
            textBox2.Text = Convert.ToString(JobExcel.Sell); 
            textBox4.Text = Convert.ToString(JobExcel.Pribil); 
            textBoxNDFL.Text = Convert.ToString(JobExcel.Ndfl); 

        }

        private void DialogVibora()
        {
            OpenFileDialog opfd = new OpenFileDialog();
            opfd.Filter = "Excel files(*.xlsx)|*.xlsx";
            if (opfd.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = opfd.FileName;

            textBox1.Text = filename;

        }

    }
}
