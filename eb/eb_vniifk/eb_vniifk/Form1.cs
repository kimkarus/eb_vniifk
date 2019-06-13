using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Security;

namespace eb_vniifk
{
    public partial class Form1 : Form
    {
        private ReadExcel readExcel;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var filePath = openFileDialog1.FileName;
                    using (FileStream fs = File.Open(filePath, FileMode.Open))
                    {
                        readExcel = new ReadExcel(filePath);

                        //Process.Start("notepad.exe", filePath);
                    }
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show("error" + ex.ToString());
                }
            }
        }
    }
}
