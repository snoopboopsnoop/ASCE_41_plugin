using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ASCE_41
{
    internal class MaterialListItem : ListView
    {
        private int CLOSED_HEIGHT = 28;
        private bool closed = true;
        private int OPEN_HEIGHT = 0;

        public MaterialListItem(int width) : base()
        {


            //this.Dock = DockStyle.Fill;
            this.Width = width;

            this.Margin = new Padding(0);

            this.View = View.Details;

            this.LabelEdit = true;

            //this.AllowColumnReorder = true;

            //this.CheckBoxes = true;

            this.FullRowSelect = true;

            this.GridLines = true;

            //this.Sorting = SortOrder.Ascending;

            this.Scrollable = false;

            ListViewItem item1 = new ListViewItem("item1", 1);
            item1.Checked = true;
            item1.SubItems.Add("1");
            item1.SubItems.Add("2");
            item1.SubItems.Add("3");

            ListViewItem item2 = new ListViewItem("item2", 1);
            item2.Checked = true;
            item2.SubItems.Add("4");
            item2.SubItems.Add("5");
            item2.SubItems.Add("6");

            ListViewItem item3 = new ListViewItem("item3", 1);
            item3.Checked = true;
            item3.SubItems.Add("7");
            item3.SubItems.Add("8");
            item3.SubItems.Add("9");

            //this.Columns.Add("Item Column", -2, HorizontalAlignment.Left);
            //this.Columns.Add("Column 2", -2, HorizontalAlignment.Left);
            //this.Columns.Add("Column 3", -2, HorizontalAlignment.Left);
            //this.Columns.Add("Column 4", -2, HorizontalAlignment.Center);

            this.Columns.Add("Material", -2, HorizontalAlignment.Left);
            this.Columns.Add("", -2, HorizontalAlignment.Left);
            this.Columns.Add("", -2, HorizontalAlignment.Left);
            this.Columns.Add("", -2, HorizontalAlignment.Center);

            this.Items.AddRange(new ListViewItem[] { item1, item2, item3 });

            //CLOSED_HEIGHT = this.TopItem.Bounds.Top;

            this.ColumnClick += column_Click;

            this.Height = CLOSED_HEIGHT;

            OPEN_HEIGHT = this.Items[0].Bounds.Height * this.Items.Count + CLOSED_HEIGHT;

            //this.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);

            //this.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void column_Click(object sender, ColumnClickEventArgs e)
        {
            if(closed)
            {
                this.Height = OPEN_HEIGHT;
                closed = false;
            }
            else
            {
                this.Height = CLOSED_HEIGHT;
                closed = true;
            }
        }
    }
}
