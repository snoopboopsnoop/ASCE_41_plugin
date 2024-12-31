using ETABSv1;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASCE_41
{
    public class cPlugin
    {

        public void Main(ref ETABSv1.cSapModel SapModel, ref ETABSv1.cPluginCallback ISapPlugin)
        {

            Main main = new Main(ref SapModel, ref ISapPlugin);
            main.Show();
        }

        public long Info(ref string Text)
        {
            Text = "Plugin by Walter Cheng";
            return 0;
        }
    }
}
