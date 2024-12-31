using ETABSv1;
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
    public partial class JointScreen : Form
    {
        private cSapModel SapModel;

        public JointScreen(ref cSapModel SapModel)
        {
            InitializeComponent();

            this.SapModel = SapModel;

            int num = 0;
            string[] names = { };
            double[] x = { };
            double[] y = { };
            double[] z = { };


            SapModel.PointObj.GetAllPoints(ref num, ref names, ref x, ref y, ref z);

            comboBox1.Items.AddRange(names);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                listBox1.Items.Clear();

                int num = 0;
                int[] types = { };
                string[] names = { };
                int[] pointNums = { };
                SapModel.PointObj.GetConnectivity(comboBox1.Text, ref num, ref types, ref names, ref pointNums);

                double x = 0;
                double y = 0;
                double z = 0;

                SapModel.PointObj.GetCoordCartesian(comboBox1.Text, ref x, ref y, ref z);
                label2.Text = "(" + x + ", " + y + ", " + z + ")";

                string groupName = "SLRS";
                int gNum = 0;
                int[] gTypes = { };
                string[] gNames = { };
                SapModel.GroupDef.GetAssignments(groupName, ref gNum, ref gTypes, ref gNames);

                string text;
                for(int i = 0; i < num; i++)
                {
                    text = names[i];
                    string point1 = "";
                    string point2 = "";
                    if (gNames.Contains(text) && types[i] == 2) ;
                    {
                        SapModel.FrameObj.GetPoints(text, ref point1, ref point2);

                        double pointX = 0;
                        double pointY = 0;
                        double pointZ = 0;

                        if (point1 == comboBox1.Text)
                        {
                            SapModel.PointObj.GetCoordCartesian(point2, ref pointX, ref pointY, ref pointZ);
                        }
                        else
                        {
                            SapModel.PointObj.GetCoordCartesian(point1, ref pointX, ref pointY, ref pointZ);
                        }
                        double length = Math.Sqrt(Math.Pow(x - pointX, 2) + Math.Pow(y - pointY, 2) + Math.Pow(z - pointZ, 2)) / 12;
                        text += " (" + length.ToString("0.##") + " ft)";

                        text += " (" + pointX + ", " + pointY + ", " + pointZ + ")";

                        if (pointX != x)
                        {
                            text += (pointX < x) ? " (left, x)" : " (right, x)";
                        }
                        else if (pointY != y)
                        {
                            text += (pointY < y) ? " (left, y)" : " (right, y)";
                        }
                        else if (pointZ != z)
                        {
                            text += (pointZ < z) ? " (bottom)" : " (top)";
                        }

                        listBox1.Items.Add(text);
                    }
                    //switch (types[i])
                    //{
                    //    case 2:
                    //        SapModel.FrameObj.GetPoints(text, ref point1, ref point2);
                    //        //text += " (Frame obj)";
                    //        break;
                    //    case 3:
                    //        text += " (Cable obj)";
                    //        break;
                    //    case 4:
                    //        text += " (Tendon obj)";
                    //        break;
                    //    case 5:
                    //        text += " (Area obj)";
                    //        break;
                    //    case 6:
                    //        text += " (Solid obj)";
                    //        break;
                    //    case 7:
                    //        text += " (Link obj)";
                    //        break;
                    //}

                    
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
