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
    public partial class MaterialList : Form
    {
        private List<Material> materials = new List<Material>();
        private Material currentMat;

        public MaterialList(ref List<Material> materials)
        {
            InitializeComponent();
            this.materials = materials;

            foreach(Material mat in materials)
            {
                int index = comboBox1.Items.Add(mat.GetName());
            }
            comboBox1.SelectedIndex = 0;
            currentMat = materials[0];
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            currentMat = materials.Find(x => x.GetName() == comboBox1.SelectedItem.ToString());
            textBox1.Text = currentMat.GetMatType().ToString();
            textBox2.Text = currentMat.GetKFactor().ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            label4.Visible = false;
            try
            {

                double kFactor = Double.Parse(textBox2.Text);
                currentMat.SetKFactor(kFactor);

                this.Cursor = Cursors.Default;

                this.Close();
            }
            catch
            {
                label4.Visible = true;
                this.Cursor = Cursors.Default;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
