using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ASCE_41
{
    public partial class ElementScreen : Form
    {
        private List<Element> elements = new List<Element>();
        private string type;

        public ElementScreen(List<Element> elements, string type)
        {
            this.elements = elements;
            InitializeComponent();

            Elem_Control_Box.Items.AddRange(new string[] { "Force Controlled", "Deformation Controlled"});
            Elem_Control_Box.SelectedIndex = 0;

            this.elements = elements.Select(element =>
            {
                if (element.GetType() == type) return element;
                else return null;
            }).ToList();

            this.elements.RemoveAll(item => item == null);

            comboBox1.Items.AddRange(this.elements.Select(element => element.GetName()).ToArray());

            comboBox1.SelectedIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox3.Text = elements[comboBox1.SelectedIndex].GetM()[0].ToString();
            textBox6.Text = elements[comboBox1.SelectedIndex].GetM()[1].ToString();

            textBox4.Text = elements[comboBox1.SelectedIndex].GetJ().ToString();

            string temp = elements[comboBox1.SelectedIndex].GetEControl();
            if (temp != "")
            {
                if (temp.Contains("Force"))
                {
                    Elem_Control_Box.SelectedIndex = 0;
                }
                else Elem_Control_Box.SelectedIndex = 1;
            }

            textBox2.Text = elements[comboBox1.SelectedIndex].GetKFactor().ToString();

            double[] factAdjust = elements[comboBox1.SelectedIndex].GetFactorAdj();

            textBox1.Text = factAdjust[0].ToString();
            textBox5.Text = factAdjust[1].ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label9.Visible = false;
            try
            {
                if(checkBox1.Checked)
                {
                    foreach(Element elem in elements)
                    {
                        elem.SetM1(double.Parse(textBox3.Text));
                        elem.SetM2(double.Parse(textBox6.Text));
                        elem.SetJ(double.Parse(textBox4.Text));
                        elem.SetControl(Elem_Control_Box.Text);
                        elem.SetKFactor(double.Parse(textBox2.Text));
                        elem.SetFactor1(double.Parse(textBox1.Text));
                        elem.SetFactor2(double.Parse(textBox5.Text));
                        elem.SetEdited(true);
                    }
                    this.Close();
                }
                else
                {
                    Element elem = elements[comboBox1.SelectedIndex];
                    elem.SetM1(double.Parse(textBox3.Text));
                    elem.SetM2(double.Parse(textBox6.Text));
                    elem.SetJ(double.Parse(textBox4.Text));
                    elem.SetControl(Elem_Control_Box.Text);
                    elem.SetKFactor(double.Parse(textBox2.Text));
                    elem.SetFactor1(double.Parse(textBox1.Text));
                    elem.SetFactor2(double.Parse(textBox5.Text));
                    elem.SetEdited(true);
                }

                this.Close();
            }
            catch(Exception)
            {
                label9.Visible = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
