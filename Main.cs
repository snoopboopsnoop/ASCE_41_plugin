using ETABSv1;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.Remoting.Messaging;

namespace ASCE_41
{

    public struct Output
    {
        public string[] SHL;
        public string[] PL;

        public string AP;

        public double dampingRatio;

        public int system;
        public double hn;

        public double c1c2;
        public double cm;

        public double[] Sxs;
        public double[] Sx1;

        public double[] Tl;


        public double sWeight;
        public double[,] iLDP;

        public string savePath;

        public List<Material> materials;
        public Dictionary<string, Element> frameElements;
        public Dictionary<string, Element> areaElements;

        public Output(
            string[] SHL,
            string[] PL,

            string AP,

            int system,
            double hn,

            double c1c2,
            double cm,

            double[] Sxs,
            double[] Sx1,

            double[] Tl,

            double sWeight,
            double[,] iLDP,

            double dampingRatio,

            List<Material> materials,
            Dictionary<string, Element> frameElements,
            Dictionary<string, Element> areaElements,

            string savePath)
        {
            this.SHL = SHL;
            this.PL = PL;

            this.AP = AP;

            this.system = system;
            this.hn = hn;

            this.c1c2 = c1c2;
            this.cm = cm;

            this.Sxs = Sxs;
            this.Sx1 = Sx1;

            this.Tl = Tl;
            this.iLDP = iLDP;

            this.dampingRatio = dampingRatio;

            this.sWeight = sWeight;

            this.materials = materials;
            this.frameElements = frameElements;
            this.areaElements = areaElements;

            this.savePath = savePath;
        }
    }

    public partial class Main : Form
    {
        private cPluginCallback _Plugin = null;
        private cSapModel _SapModel = null;

        private Dictionary<string, Element> frameElements = new Dictionary<string, Element>();
        private Dictionary<string, Element> areaElements = new Dictionary<string, Element>();
        private List<Element> selectedElems = new List<Element>();

        private List<string> console = new List<string>();
        private List<Material> materials = new List<Material>();

        private string[] SHLOptions = { "BSE-1E", "BSE-2E", "BSE-1N", "BSE-2N" };
        private string[] PLOptions = { "S-1(Immediate Occupancy, IO)", "S-2(Damage Control, DC)", "S-3(Life Safety, LS)", "S-4(Limited Safety, LimS)", "S-5(Collapse Prevention, CP)" };
        private string[] APOptions = { "LSP", "LDP", "HLSP", "NLDP" };

        private string[] Systems = { "Steel Moment-Resisting", "Concrete Moment-Resisting", "Steel Eccentrically Braced", "Other" };
        private int system = -1;

        private string[] headers = { "k", "m", "J", "Element Control", "X" };
        //private string[] headers = { "Knowledge Factor (k)", "Component capacity modification factor (m)", "Force-delivery reduction factor (J)", "Element Control", "Factor adjusting for \"PL\" (X)" };

        private string[] SHL = { "", "" };
        private string[] PL = { "", "" };

        private string AP = "";

        private double period = 0.0;
        private double ct = 0.0;
        private double hn = 0.0;
        private double beta = 0.0;

        private double dampingRatio = 0.05;

        private double c1c2 = 0.0;
        private double cm = 0.0;
        private double b1 = 0.0;

        private double[] Sxs = { 0.0, 0.0 };
        private double[] Sx1 = { 0.0, 0.0 };

        private double[] Tl = { 0.0, 0.0 };
        private double[] T0 = { 0.0, 0.0 };
        private double[] Ts = { 0.0, 0.0 };

        private double sWeight = 0.0;

        // per floor
        private double[] Sa = { 0.0, 0.0 };
        private double[] vw = { 0.0, 0.0 };
        private double[] VSHL = { 0.0, 0.0 };

        private double[] scaleFactor = { 0.0, 0.0 };
        private double[,] iLDP = { { 0.0, 0.0 }, { 0.0 , 0.0} };
        private double[] LSP85 = { 0.0, 0.0 };
        private double[,] finalFactor = { { 0.0, 0.0 }, { 0.0, 0.0 } };

        private string savePath = "";
        private object templatePath = Path.Combine(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName, "src/template.docx");

        private MenuStrip menu = new MenuStrip();

        public Main(ref cSapModel SapModel, ref cPluginCallback Plugin)
        {
            _Plugin = Plugin;
            _SapModel = SapModel;

            InitializeComponent();

            Controls.Add(menu);
            ToolStripMenuItem fileMenu = new ToolStripMenuItem("File");
            ToolStripMenuItem openMenu = new ToolStripMenuItem("Open...", null, new EventHandler(open_Click));
            ToolStripMenuItem saveMenu = new ToolStripMenuItem("Save", null, new EventHandler(save_Click));
            ToolStripMenuItem saveAsMenu = new ToolStripMenuItem("Save As...", null, new EventHandler(saveAs_Click));

            fileMenu.DropDownItems.Add(openMenu);
            fileMenu.DropDownItems.Add(saveMenu);
            fileMenu.DropDownItems.Add(saveAsMenu);

            menu.Items.Add(fileMenu);

            ToolStripMenuItem editMenu = new ToolStripMenuItem("Edit");
            ToolStripMenuItem matMenu = new ToolStripMenuItem("Set Material K Factors", null, new EventHandler(EditMaterials));

            editMenu.DropDownItems.Add(matMenu);

            menu.Items.Add(editMenu);

            ToolStripMenuItem viewMenu = new ToolStripMenuItem("View");
            ToolStripMenuItem viewFrameElemMenu = new ToolStripMenuItem("Selected Frame Elements", null, new EventHandler(ViewFrameElements));
            ToolStripMenuItem viewAreaElemMenu = new ToolStripMenuItem("Selected Area Elements", null, new EventHandler(ViewAreaElements));

            viewMenu.DropDownItems.Add(viewFrameElemMenu);
            viewMenu.DropDownItems.Add(viewAreaElemMenu);

            menu.Items.Add(viewMenu);

            ToolStripMenuItem jointMenu = new ToolStripMenuItem("Joints");
            ToolStripMenuItem lookupJointMenu = new ToolStripMenuItem("Joint Lookup", null, new EventHandler(lookup_Click));
            ToolStripMenuItem exportJointMenu = new ToolStripMenuItem("Export Selected Joints", null, new EventHandler(exportJoint_Click));

            jointMenu.DropDownItems.Add(lookupJointMenu);
            jointMenu.DropDownItems.Add(exportJointMenu);

            menu.Items.Add(jointMenu);

            ToolStripMenuItem exportMenu = new ToolStripMenuItem("Export");
            ToolStripMenuItem exportFrameMenu = new ToolStripMenuItem("Export Selected Frames", null, new EventHandler(ExportFrames));
            ToolStripMenuItem exportAreaMenu = new ToolStripMenuItem("Export All Area Elements", null, new EventHandler(ExportAreas));

            exportMenu.DropDownItems.Add(exportFrameMenu);
            exportMenu.DropDownItems.Add(exportAreaMenu);

            menu.Items.Add(exportMenu);

            //label7.Text = "";
            SHL_Level_1_Box.Items.AddRange(SHLOptions);
            SHL_Level_2_Box.Items.AddRange(SHLOptions);

            SHL_Level_1_Box.DropDownStyle = ComboBoxStyle.DropDownList;
            SHL_Level_2_Box.DropDownStyle = ComboBoxStyle.DropDownList;

            PL_Level_1_Box.Items.AddRange(PLOptions);
            PL_Level_2_Box.Items.AddRange(PLOptions);

            PL_Level_1_Box.DropDownStyle = ComboBoxStyle.DropDownList;
            PL_Level_2_Box.DropDownStyle = ComboBoxStyle.DropDownList;

            dRatioBox.Text = dampingRatio.ToString("0.00");

            AP_Box.Items.AddRange(APOptions);
            AP_Box.DropDownStyle = ComboBoxStyle.DropDownList;

            FSystemBox.Items.AddRange(Systems);

            //Period_Box.Parent = this;

            //Period_Box.LostFocus += control_Leave;

            this.MouseClick += Main_MouseDown;

            //button2.Visible = false;
            //label7.Visible = false;

            this.FormClosing += onClose;

            GetMaterials();
        }

        private void onClose(object sender, EventArgs e)
        {
            MessageScreen msg = new MessageScreen("Save before closing?", "Yes", "No");
            if(msg.ShowDialog() == DialogResult.OK)
            {
                Save();
            }
        }

        private void open_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Json files (*.json)|*.json|Text files (*.txt)|*.txt";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                LoadData(openDialog.FileName);
            }
        }
        private void save_Click(object sender, EventArgs e)
        {
            Save();
        }
        private void saveAs_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void lookup_Click(object sender, EventArgs e)
        {
            JointScreen joints = new JointScreen(ref _SapModel);
            joints.Show();
        }
        private void exportJoint_Click(object sender, EventArgs e)
        {
            Excel.Application app;
            try
            {
                app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;
            }
            catch(Exception ex)
            {
                app = new Excel.Application();
            }

            app.Visible = true;

            Workbook workbook = app.Workbooks.Add();

            Excel.Worksheet worksheet = workbook.ActiveSheet;

            app.ActiveWindow.Activate();

            try
            {
                Excel.Range header = worksheet.UsedRange.Rows[1].EntireRow;
                header.Interior.Color = System.Drawing.Color.Yellow;
                header.Borders.LineStyle = worksheet.Range["A2", "A2"].Borders.LineStyle;

                int numNames = 0;
                string[] pointNames = { };
                double[] X = { };
                double[] Y = { };
                double[] Z = { };

                _SapModel.PointObj.GetAllPoints(ref numNames, ref pointNames, ref X, ref Y, ref Z);

                int row = 2;

                for (int i = 0; i < numNames; ++i)
                {
                    string elemName = pointNames[i];

                    bool selected = false;
                    _SapModel.PointObj.GetSelected(elemName, ref selected);

                    if (!selected) continue;

                    string label = "";
                    string story = "";
                    _SapModel.PointObj.GetLabelFromName(elemName, ref label, ref story);

                    worksheet.Cells[row, 1] = story;
                    worksheet.Cells[row, 2] = label;
                    worksheet.Cells[row, 3] = elemName;
                    

                    int num = 0;
                    int[] types = { };
                    string[] names = { };
                    int[] pointNums = { };
                    _SapModel.PointObj.GetConnectivity(elemName, ref num, ref types, ref names, ref pointNums);

                    double x = 0;
                    double y = 0;
                    double z = 0;

                    _SapModel.PointObj.GetCoordCartesian(elemName, ref x, ref y, ref z);
                    worksheet.Cells[row, 4] = "(" + x + ", " + y + ", " + z + ")";

                    for (int j = 0; j < num; j++)
                    {

                        string point1 = "";
                        string point2 = "";

                        _SapModel.FrameObj.GetPoints(names[j], ref point1, ref point2);

                        double pointX = 0;
                        double pointY = 0;
                        double pointZ = 0;

                        if (point1 == elemName)
                        {
                            _SapModel.PointObj.GetCoordCartesian(point2, ref pointX, ref pointY, ref pointZ);
                        }
                        else
                        {
                            _SapModel.PointObj.GetCoordCartesian(point1, ref pointX, ref pointY, ref pointZ);
                        }


                        string section = "";
                        string sAuto = "";
                        _SapModel.FrameObj.GetSection(names[j], ref section, ref sAuto);

                        string frameLabel = "";
                        string frameStory = "";
                        _SapModel.FrameObj.GetLabelFromName(names[j], ref frameLabel, ref frameStory);

                        if (pointX != x)
                        {
                            if(pointX < x)
                            {
                                //left x
                                worksheet.Cells[row, 6] = section;
                                worksheet.Cells[row, 7] = label;
                                worksheet.Cells[row, 8] = names[j];
                            }
                            else
                            {
                                //right x
                                worksheet.Cells[row, 22] = section;
                                worksheet.Cells[row, 23] = label;
                                worksheet.Cells[row, 24] = names[j];
                            }
                        }
                        else if (pointY != y)
                        {
                            if (pointY < y)
                            {
                                // left y
                                worksheet.Cells[row, 10] = section;
                                worksheet.Cells[row, 11] = label;
                                worksheet.Cells[row, 12] = names[j];
                            }
                            else
                            {
                                // right y
                                worksheet.Cells[row, 26] = section;
                                worksheet.Cells[row, 27] = label;
                                worksheet.Cells[row, 28] = names[j]; 
                            }
                        }
                        else if (pointZ != z)
                        {
                            if (pointZ < z)
                            {
                                // bottom z
                                worksheet.Cells[row, 18] = section;
                                worksheet.Cells[row, 19] = label;
                                worksheet.Cells[row, 20] = names[j];
                            }
                            else
                            {
                                // top z
                                worksheet.Cells[row, 14] = section;
                                worksheet.Cells[row, 15] = label;
                                worksheet.Cells[row, 16] = names[j];
                            }
                        }

                    }
                    row++;
                }

                worksheet.Cells[1, 1] = "Story";
                worksheet.Cells[1, 2] = "Label";
                worksheet.Cells[1, 3] = "Unique Name";

                worksheet.Cells[1, 6] = "Beam Left, X";
                worksheet.Cells[1, 7] = "Label";
                worksheet.Cells[1, 8] = "Unique Name";

                worksheet.Cells[1, 10] = "Beam Left, Y";
                worksheet.Cells[1, 11] = "Label";
                worksheet.Cells[1, 12] = "Unique Name";

                worksheet.Cells[1, 14] = "Column Above";
                worksheet.Cells[1, 15] = "Label";
                worksheet.Cells[1, 16] = "Unique Name";

                worksheet.Cells[1, 18] = "Column Below";
                worksheet.Cells[1, 19] = "Label";
                worksheet.Cells[1, 20] = "Unique Name";

                worksheet.Cells[1, 22] = "Beam Right, X";
                worksheet.Cells[1, 23] = "Label";
                worksheet.Cells[1, 24] = "Unique Name";

                worksheet.Cells[1, 26] = "Beam Right, Y";
                worksheet.Cells[1, 27] = "Label";
                worksheet.Cells[1, 28] = "Unique Name";

                worksheet.Range["A1", "AB" + row].Columns.AutoFit();
            }
            catch (Exception ex)
            {
                new MessageScreen("The Joints could not be exported:\n\n" + ex.Message).Show();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
        }

        private bool Save()
        {
            UpdateElements();

            if(savePath == "")
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Json files (*.json)|*.json|Text files (*.txt)|*.txt";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    if (savePath == "") savePath = saveDialog.FileName;
                }
                else return false;
            }

            this.Cursor = Cursors.WaitCursor;

            JsonSerializer serializer = new JsonSerializer();
            serializer.TypeNameHandling = TypeNameHandling.Auto;
            serializer.ReferenceLoopHandling = ReferenceLoopHandling.Serialize;

            using (StreamWriter sw = new StreamWriter(savePath))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                serializer.Serialize(writer, new Output(SHL, PL, AP, system, hn, c1c2, cm, Sxs, Sx1, Tl, sWeight, iLDP, dampingRatio, materials, frameElements, areaElements, savePath));
            }
            this.Cursor = Cursors.Default;
            return true;
        }

        private void LoadData(string path)
        {
            try
            {
                using (StreamReader file = File.OpenText(path))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.TypeNameHandling = TypeNameHandling.Auto;
                    serializer.NullValueHandling = NullValueHandling.Ignore;

                    Output input = (Output)serializer.Deserialize(file, typeof(Output));

                    this.SHL = input.SHL;
                    SHL_Level_1_Box.Text = this.SHL[0];
                    SHL_Level_2_Box.Text = this.SHL[1];

                    this.PL = input.PL;
                    PL_Level_1_Box.Text = this.PL[0];
                    PL_Level_2_Box.Text = this.PL[1];

                    this.AP = input.AP;
                    AP_Box.Text = this.AP;

                    if (input.system != -1)
                    {
                        FSystemBox.SelectedIndex = input.system;
                    }

                    this.hn = input.hn;
                    HnBox.Text = this.hn.ToString("0.00");

                    this.c1c2 = input.c1c2;
                    cBox.Text = this.c1c2.ToString("0.00");
                    this.cm = input.cm;
                    CmBox.Text = this.cm.ToString("0.00");

                    this.Sxs = input.Sxs;
                    Sxs1Box.Text = this.Sxs[0].ToString("0.000");
                    Sxs2Box.Text = this.Sxs[1].ToString("0.000");
                    this.Sx1 = input.Sx1;
                    Sx1Box.Text = this.Sx1[0].ToString("0.000");
                    Sx2Box.Text = this.Sx1[1].ToString("0.000");


                    this.Tl = input.Tl;
                    Tl1Box.Text = this.Tl[0].ToString("0.000");
                    Tl2Box.Text = this.Tl[1].ToString("0.000");

                    this.sWeight = input.sWeight;
                    SWeightBox.Text = this.sWeight.ToString("0.0");

                    this.iLDP = input.iLDP;
                    iLDPx1Box.Text = this.iLDP[0, 0].ToString("0.00");
                    iLDPy1Box.Text = this.iLDP[0, 1].ToString("0.00");
                    iLDPx2Box.Text = this.iLDP[1, 0].ToString("0.00");
                    iLDPy2Box.Text = this.iLDP[1, 1].ToString("0.00");

                    this.dampingRatio = input.dampingRatio;
                    dRatioBox.Text = this.dampingRatio.ToString("0.00");

                    this.savePath = input.savePath;

                    this.materials = input.materials;
                    this.frameElements = input.frameElements;
                    this.areaElements = input.areaElements;

                    GetMaterials();
                    UpdateElements();
                }
            }
            catch (Exception e)
            {
                new MessageScreen(e.Message).Show();
            }
        }

        private void Main_MouseDown(object sender, MouseEventArgs e)
        {
            //console.Add("clicky " + click);
            this.ActiveControl = null;
        }

        // has different 
        private void control_Leave(object sender, EventArgs e)
        {
            (this.Parent as Form).ActiveControl = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (console.Count > 1)
            {
                console.RemoveAt(0);
                //label7.Text = console[0];

            }
        }

        private void EditMaterials(object sender, EventArgs e)
        {
            string[] names = { };
            int i = 0;

            _SapModel.PropMaterial.GetNameList(ref i, ref names);

            MaterialList matList = new MaterialList(ref materials);
            matList.Show();

            UpdateElements();
        }

        private void GetMaterials()
        {
            materials.Clear();

            string[] names = { };
            int i = 0;

            _SapModel.PropMaterial.GetNameList(ref i, ref names);
            foreach (string name in names)
            {
                if(!materials.Any(mat => mat.GetName() == name))
                {
                    eMatType type = new eMatType();
                    int color = 0;
                    string notes = "";
                    string GUID = "";
                    _SapModel.PropMaterial.GetMaterial(name, ref type, ref color, ref notes, ref GUID);

                    materials.Add(new Material(name, type, color, notes, GUID));
                }
            }
        }

        private void UpdateElements()
        {
            selectedElems.Clear();

            Dictionary<string, Element> frameList = new Dictionary<string, Element>();

            string[] names = { };
            string[] propNames = { };
            string[] storyNames = { };

            string[] PointName1 = {};
            string[] PointName2 = {};
            double[] Point1X = {};
            double[] Point1Y = {};
            double[] Point1Z = {};
            double[] Point2X = {};
            double[] Point2Y = {};
            double[] Point2Z = {};
            double[] Angle = {};
            double[] Offset1X = {};
            double[] Offset2X = {};
            double[] Offset1Y = {};
            double[] Offset2Y = {};
            double[] Offset1Z = {};
            double[] Offset2Z = {};

            int[] CardinalPoint = { };

            int i = 0;

            _SapModel.FrameObj.GetAllFrames(ref i, ref names, ref propNames, ref storyNames, ref PointName1,
            ref  PointName2,
            ref Point1X,
            ref Point1Y,
            ref Point1Z,
            ref Point2X,
            ref Point2Y,
            ref Point2Z,
            ref Angle,
            ref Offset1X,
            ref Offset2X,
            ref Offset1Y,
            ref Offset2Y,
            ref Offset1Z,
            ref Offset2Z,
            ref CardinalPoint);

            for(int j = 0; j < i; ++j)
            {
                string name = names[j];

                bool selected = false;

                eMatType type = new eMatType();
                int color = 0;
                string notes = "";
                string GUID = "";

                string matProp = "";
                _SapModel.PropFrame.GetMaterial(propNames[j], ref matProp);

                if(matProp == "") continue;

                _SapModel.PropMaterial.GetMaterial(matProp, ref type, ref color, ref notes, ref GUID);

                //console.Add("name" + name + "\n" + "propName: " + propNames[j] + "\nstory: " + storyNames[j]);

                Element newElem;

                //console.Add(type.ToString());

                if(frameElements.ContainsKey(name))
                {
                    // if k not manually changed or material of element has been changed,
                    // change k to the default kFactor of the material
                    if (!frameElements[name].KEdited() || frameElements[name].GetName() != matProp)
                    {
                        frameElements[name].SetKFactor(materials.Find(x => x.GetName() == matProp).GetKFactor());
                        frameElements[name].SetPropName(propNames[j]);
                        frameElements[name].SetMatType(type);
                        frameElements[name].SetEdited(false);
                    }
                    newElem = frameElements[name];
                }
                else
                {
                    newElem = new Element(name, propNames[j], "frame", type, materials.Find(x => x.GetName() == matProp).GetKFactor());
                }

                _SapModel.FrameObj.GetSelected(name, ref selected);
                if (selected)
                {
                    selectedElems.Add(newElem);
                }

                frameList.Add(name, newElem);
            }

            this.frameElements = frameList;

            Dictionary<string, Element> areaList = new Dictionary<string, Element>();

            eAreaDesignOrientation[] designOrientations = { };
            int boundaryPts = 0;
            int[] pointDelimeter = { };
            string[] pointNames = { };
            double[] pointX = { };
            double[] pointY = { };
            double[] pointZ = { };

            _SapModel.AreaObj.GetAllAreas(ref i, ref names, ref designOrientations, ref boundaryPts, ref pointDelimeter, ref pointNames, ref pointX, ref pointY, ref pointZ);

            for (int j = 0; j < i; ++j)
            {
                string name = names[j];

                bool selected = false;

                string propName = "";
                _SapModel.AreaObj.GetProperty(name, ref propName);

                eWallPropType wallPropType = new eWallPropType();
                eShellType shellType = new eShellType();
                string matProp = "";
                double thickness = 0;
                int color = 0;
                string notes = "";
                string GUID = "";
                
                _SapModel.PropArea.GetWall(propName, ref wallPropType, ref shellType, ref matProp, ref thickness, ref color, ref notes, ref GUID);

                if (matProp == "") continue;

                eMatType type = new eMatType();

                _SapModel.PropMaterial.GetMaterial(matProp, ref type, ref color, ref notes, ref GUID);


                string pierName = "";
                _SapModel.AreaObj.GetPier(name, ref pierName);

                string spandrelName = "";
                _SapModel.AreaObj.GetSpandrel(name, ref spandrelName);

                //console.Add("name" + name + "\n" + "propName: " + propNames[j] + "\nstory: " + storyNames[j]);

                Element newElem;

                //console.Add(type.ToString());

                if (areaElements.ContainsKey(name))
                {
                    // if k not manually changed or material of element has been changed,
                    // change k to the default kFactor of the material
                    if (!areaElements[name].KEdited() || areaElements[name].GetName() != matProp)
                    {
                        areaElements[name].SetKFactor(materials.Find(x => x.GetName() == matProp).GetKFactor());
                        areaElements[name].SetPropName(propName);
                        areaElements[name].SetMatType(type);
                        areaElements[name].SetEdited(false);
                    }
                    newElem = areaElements[name];
                }
                else
                {
                    newElem = new Element(name, propName, "area", pierName, spandrelName, type, materials.Find(x => x.GetName() == matProp).GetKFactor());
                }

                _SapModel.AreaObj.GetSelected(name, ref selected);
                if (selected)
                {
                    selectedElems.Add(newElem);
                }

                areaList.Add(name, newElem);
            }

            // replace old list with new one to get rid of any deleted elements
            this.areaElements = areaList;
        }

        private void SHL_Level_1_Box_SelectedIndexChanged(object sender, EventArgs e)
        {
            SHL[0] = SHL_Level_1_Box.SelectedItem.ToString();
            Sa1Label.Text = "Level 1: Sₐ_" + SHL[0] + " = ";
        }

        private void SHL_Level_2_Box_SelectedIndexChanged(object sender, EventArgs e)
        {
            SHL[1] = SHL_Level_2_Box.SelectedItem.ToString();
            Sa2Label.Text = "Level 2: Sₐ_" + SHL[1] + " = ";
        }

        private void PL_Level_1_Box_SelectedIndexChanged(object sender, EventArgs e)
        {
            PL[0] = PL_Level_1_Box.SelectedItem.ToString();
        }

        private void PL_Level_2_Box_SelectedIndexChanged(object sender, EventArgs e)
        {
            PL[1] = PL_Level_2_Box.SelectedItem.ToString();
        }

        private void AP_Box_SelectedIndexChanged(object sender, EventArgs e)
        {
            AP = AP_Box.SelectedItem.ToString();
            if(AP == "LSP")
            {
                LDPGroup.Visible = false;
            }
            else
            {
                LDPGroup.Visible = true;
            }
        }

        private void Period_Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(Period_Box.Text);
                this.period = num;
            }
            catch(Exception) { }
        }

        private void cBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(cBox.Text);
                this.c1c2 = num;
                SetVW();
                SetScaleFactor();
            }
            catch (Exception) { }
        }

        private void HnBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(HnBox.Text);
                this.hn = num;
                SetT();
            }
            catch (Exception) {}
        }

        private void Tl1Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(Tl1Box.Text);
                this.Tl[0] = num;
                SetSa();
            }
            catch (Exception) { }
        }
        private void Tl2Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(Tl2Box.Text);
                this.Tl[1] = num;
                SetSa();
            }
            catch (Exception) { }
        }

        private void ViewFrameElements(object sender, EventArgs e)
        {

            UpdateElements();

            if (selectedElems.Count != 0)
            {
                ElementScreen elementScreen = new ElementScreen(selectedElems, "frame");
                elementScreen.Show();
            }
            else
            {
                MessageScreen message = new MessageScreen("No elements selected");
            }
        }
        private void ViewAreaElements(object sender, EventArgs e)
        {

            UpdateElements();

            if (selectedElems.Count != 0)
            {
                ElementScreen elementScreen = new ElementScreen(selectedElems, "area");
                elementScreen.Show();
            }
            else
            {
                MessageScreen message = new MessageScreen("No elements selected");
            }
        }

        private int WriteFrameSheet(Worksheet dataSheet, Worksheet writeSheet, ProgressBar progressScreen, Excel.Application app)
        {
            int startRow = 4;
            int writeRow = writeSheet.UsedRange.Rows.Count + 1;

            try
            {
                int rows = dataSheet.UsedRange.Rows.Count;

                //List<string> storyList = ((object[,])dataSheet.Range["A4" + ":A" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToString(o));
                List<string> nameList = ((object[,])dataSheet.Range["C" + startRow + ":C" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToString(o));
                List<double> pList = ((object[,])dataSheet.Range["F" + startRow + ":F" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> v2List = ((object[,])dataSheet.Range["G" + startRow + ":G" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> v3List = ((object[,])dataSheet.Range["H" + startRow + ":H" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> tList = ((object[,])dataSheet.Range["I" + startRow + ":I" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> m2List = ((object[,])dataSheet.Range["J" + startRow + ":J" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> m3List = ((object[,])dataSheet.Range["K" + startRow + ":K" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                
                for (int i = 0; i < nameList.Count; ++i)
                {
                    string elemName = nameList[i];

                    int last = nameList.LastIndexOf(elemName);

                    int range = last - i + 1;

                    if (selectedElems.Exists(elem => elem.GetName() == elemName))
                    {
                        writeSheet.Cells[writeRow, 1] = dataSheet.Cells[i + 4, 1].Value;
                        writeSheet.Cells[writeRow, 2] = dataSheet.Cells[i + 4, 2].Value;
                        writeSheet.Cells[writeRow, 3] = elemName;
                        writeSheet.Cells[writeRow, 4] = frameElements[elemName].GetPropName();
                        double P = pList.GetRange(i, range).Min();
                        writeSheet.Cells[writeRow, 5] = (P > 0) ? 0 : Math.Abs(P);
                        writeSheet.Cells[writeRow, 6] = v2List.GetRange(i, range).OrderByDescending(Math.Abs).First();
                        writeSheet.Cells[writeRow, 7] = v3List.GetRange(i, range).OrderByDescending(Math.Abs).First();
                        writeSheet.Cells[writeRow, 8] = tList.GetRange(i, range).OrderByDescending(Math.Abs).First();
                        writeSheet.Cells[writeRow, 9] = m2List.GetRange(i, range).OrderByDescending(Math.Abs).First();
                        writeSheet.Cells[writeRow, 10] = m3List.GetRange(i, range).OrderByDescending(Math.Abs).First();

                        writeRow++;

                    }

                    progressScreen.Increment(range);
                    i = last;
                }
                writeSheet.Cells[1, 1] = "bobr";
            }

            catch(Exception e)
            {
                //console.Add(e.Message);
                progressScreen.Close();

                return 0;
            }

            return 1;
        }

        private void ExportFrames(object sender, EventArgs e)
        {
            if (Save())
            {
                MessageScreen confirm = new MessageScreen("Please make sure Excel is closed before continuing.");
                if (confirm.ShowDialog() == DialogResult.OK)
                {
                    foreach (Process clsProcess in Process.GetProcessesByName("EXCEL"))
                    {
                        clsProcess.Kill();
                    }

                    UpdateElements();

                    if(selectedElems.Count == 0)
                    {
                        MessageScreen noSelected = new MessageScreen("No Frame Elements selected. This will generate a blank Excel file.");
                        if (noSelected.ShowDialog() == DialogResult.Cancel) return;
                    }

                    this.Cursor = Cursors.WaitCursor;

                    int numTables = 0;
                    string[] tableKeys = { };
                    string[] tableNames = { };
                    int[] import = { };
                    bool[] empty = { };

                    _SapModel.DatabaseTables.GetAllTables(ref numTables, ref tableKeys, ref tableNames, ref import, ref empty);

                    bool found = false;

                    int beamTable = Array.IndexOf(tableNames, "Design Forces - Beams");
                    int columnTable = Array.IndexOf(tableNames, "Design Forces - Columns");

                    if (!empty[beamTable] && !empty[columnTable])
                    {
                        found = true;

                        string[] tableKeyList = { tableKeys[beamTable], tableKeys[columnTable] };
                        _SapModel.DatabaseTables.ShowTablesInExcel(ref tableKeyList, 0);

                            
                        Excel.Application app =
                            System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;

                        app.Visible = false;

                        Workbook workbook = app.ActiveWorkbook;

                        Worksheet beamSheet = workbook.Worksheets.Item[1] as Worksheet;
                        Worksheet columnSheet = workbook.Worksheets.Item[2] as Worksheet;
                        Worksheet newSheet = workbook.Worksheets.Add();

                        ProgressBar progressScreen = new ProgressBar(beamSheet.UsedRange.Rows.Count + columnSheet.UsedRange.Rows.Count - 6);

                        try
                        {
                            newSheet.Name = "Design Forces - ASCE 41";

                            bool level1 = beamSheet.Cells[4, 4].Columns.Find("PO1") != null;

                            Excel.Range header = beamSheet.Range["A1", "K3"];
                            header.Copy();
                            Excel.Range newHeader = newSheet.Range["A1", "K3"];
                            newHeader.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone);

                            newSheet.Columns["E"].Delete();

                            newSheet.Cells[2, 4].Value = "Section Name";
                            
                            progressScreen.Show();

                            WriteFrameSheet(beamSheet, newSheet, progressScreen, app);

                            WriteFrameSheet(columnSheet, newSheet, progressScreen, app);

                            newSheet.Cells[1, 14] = "Table: Steel Frame Design Summary - ASCE 41 Analysis";

                            Excel.Range titleHeader = newSheet.Range["A1", "A1"].EntireRow;

                            titleHeader.Interior.Color = newSheet.Range["A1", "A1"].Interior.Color;
                            titleHeader.Font.Bold = true;

                            newSheet.Range["A2", "A2"].EntireRow.Interior.Color = newSheet.Range["A2", "A2"].Interior.Color;
                            newSheet.Range["A2", "A2"].EntireRow.Borders.LineStyle = newSheet.Range["A3", "A3"].Borders.LineStyle;
                            newSheet.Range["A3", "A3"].EntireRow.Interior.Color = newSheet.Range["A3", "A3"].Interior.Color;
                            newSheet.Range["A3", "A3"].EntireRow.Borders.LineStyle = newSheet.Range["A2", "A2"].Borders.LineStyle;
                            newSheet.Range["A2", "A2"].EntireRow.Font.Bold = true;

                            for (int j = 0; j < headers.Count(); ++j)
                            {
                                newSheet.Cells[2, 14 + j] = headers[j];
                            }


                            newSheet.Range["M2", "Q2"].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            int rowCount = newSheet.UsedRange.Rows.Count;

                            for (int j = 4; j <= rowCount; ++j)
                            {
                                Element elem = frameElements[newSheet.Cells[j, 3].Value.ToString()];

                                newSheet.Cells[j, 14] = elem.GetKFactor();
                                newSheet.Cells[j, 15] = (level1) ? elem.GetM()[0] : elem.GetM()[1];
                                newSheet.Cells[j, 16] = elem.GetJ();
                                newSheet.Cells[j, 17] = elem.GetEControl();
                                newSheet.Cells[j, 18] = (level1) ? elem.GetFactorAdj()[0] : elem.GetFactorAdj()[1];
                            }

                            string end = "Q" + newSheet.UsedRange.Rows.Count;
                            newSheet.Range["A2", end].Columns.AutoFit();

                            //console.Add(progressScreen.Value() + "/" + progressScreen.Maximum() + "\n" + selectedElems.Count + " selected elements");
                            app.Visible = true;
                            progressScreen.Close();
                        }
                        catch (Exception ex)
                        {
                            //console.Add(ex.Message);
                            new MessageScreen("An error occurred. The Excel sheet could not be generated.").Show();
                            progressScreen.Close();
                        }
                            

                        // interop cleanup
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(beamSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(columnSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(newSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                    }

                    if (!found)
                    {
                        new MessageScreen("Could not find Steel Frame Analysis Summary table - Has analysis been run?").Show();
                    }

                    this.Cursor = Cursors.Default;
                }
            }
        }

        private int WriteAreaSheet(Worksheet dataSheet, Worksheet writeSheet, ProgressBar progressScreen, Excel.Application app, bool level1)
        {
            int startRow = writeSheet.UsedRange.Rows.Count + 1;
            int writeRow = startRow;

            try
            {
                int rows = dataSheet.UsedRange.Rows.Count;

                List<string> nameList = ((object[,])dataSheet.Range["B" + startRow + ":B" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToString(o));
                List<double> pList = ((object[,])dataSheet.Range["E" + startRow + ":E" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> v2List = ((object[,])dataSheet.Range["F" + startRow + ":F" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> v3List = ((object[,])dataSheet.Range["G" + startRow + ":G" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> tList = ((object[,])dataSheet.Range["H" + startRow + ":H" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> m2List = ((object[,])dataSheet.Range["I" + startRow + ":I" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));
                List<double> m3List = ((object[,])dataSheet.Range["J" + startRow + ":J" + (rows)].Value2).Cast<object>().ToList().ConvertAll(o => Convert.ToDouble(o));

                for (int i = 0; i < nameList.Count; ++i)
                {
                    string name = nameList[i];

                    int last = nameList.LastIndexOf(name);

                    int range = last - i + 1;

                    writeSheet.Cells[writeRow, 1] = dataSheet.Cells[i + 4, 1].Value;
                    writeSheet.Cells[writeRow, 2] = name;
                    //writeSheet.Cells[writeRow, 4] = areaElements[elemName].GetPropName();
                    double P = pList.GetRange(i, range).Max();
                    writeSheet.Cells[writeRow, 4] = (P < 0) ? 0 : P;
                    writeSheet.Cells[writeRow, 5] = v2List.GetRange(i, range).OrderByDescending(Math.Abs).First();
                    writeSheet.Cells[writeRow, 6] = v3List.GetRange(i, range).OrderByDescending(Math.Abs).First();
                    writeSheet.Cells[writeRow, 7] = tList.GetRange(i, range).OrderByDescending(Math.Abs).First();
                    writeSheet.Cells[writeRow, 8] = m2List.GetRange(i, range).OrderByDescending(Math.Abs).First();
                    writeSheet.Cells[writeRow, 9] = m3List.GetRange(i, range).OrderByDescending(Math.Abs).First();

                    List<Element> elems = areaElements.Select(elem =>
                    {
                        if (elem.Value.GetPierName() == name || elem.Value.GetSpandrelName() == name)
                        {
                            return elem.Value;
                        }
                        else return null;
                    }).ToList();

                    elems.RemoveAll(item => item == null);

                    // check that all k factors are the same
                    bool error = !(elems.All(elem =>
                    {
                        return elem.GetKFactor() == elems[0].GetKFactor();
                    }));

                    if (error)
                    {
                        writeSheet.Cells[writeRow, 13].Font.Color = Color.Red;
                        writeSheet.Cells[writeRow, 13] = "Error: At least one element has a mismatching K factor";
                    }
                    else
                    {
                        writeSheet.Cells[writeRow, 13] = elems[0].GetKFactor();
                    }

                    // check that all M factors are the same
                    error = !(elems.All(elem =>
                    {
                        if (level1)
                        {
                            return elem.GetM()[0] == elems[0].GetM()[0];
                        }
                        else return elem.GetM()[1] == elems[0].GetM()[1];
                    }));

                    if (error)
                    {
                        writeSheet.Cells[writeRow, 14].Font.Color = Color.Red;
                        writeSheet.Cells[writeRow, 14] = "Error: At least one element has a mismatching M factor";
                    }
                    else
                    {
                        writeSheet.Cells[writeRow, 14] = (level1) ? elems[0].GetM()[0] : elems[0].GetM()[1];
                    }

                    // check that all J factors are the same
                    error = !(elems.All(elem =>
                    {
                        return elem.GetJ() == elems[0].GetJ();
                    }));

                    if (error)
                    {
                        writeSheet.Cells[writeRow, 15].Font.Color = Color.Red;
                        writeSheet.Cells[writeRow, 15] = "Error: At least one element has a mismatching J factor";
                    }
                    else
                    {
                        writeSheet.Cells[writeRow, 15] = elems[0].GetJ();
                    }

                    // check that all EControl are the same
                    error = !(elems.All(elem =>
                    {
                        return elem.GetEControl() == elems[0].GetEControl();
                    }));

                    if (error)
                    {
                        writeSheet.Cells[writeRow, 16].Font.Color = Color.Red;
                        writeSheet.Cells[writeRow, 16] = "Error: At least one element has a mismatching EControl factor";
                    }
                    else
                    {
                        writeSheet.Cells[writeRow, 16] = elems[0].GetEControl();
                    }

                    // check that all factoradj are the same
                    error = !(elems.All(elem =>
                    {
                        if (level1)
                        {
                            return elem.GetFactorAdj()[0] == elems[0].GetFactorAdj()[0];
                        }
                        else return elem.GetFactorAdj()[1] == elems[0].GetFactorAdj()[1];
                    }));

                    if (error)
                    {
                        writeSheet.Cells[writeRow, 17].Font.Color = Color.Red;
                        writeSheet.Cells[writeRow, 17] = "Error: At least one element has a mismatching M factor";
                    }
                    else
                    {
                        writeSheet.Cells[writeRow, 17] = (level1) ? elems[0].GetFactorAdj()[0] : elems[0].GetFactorAdj()[1];
                    }

                    string elemOut = String.Join(", ", elems.Select(elem => elem.GetName()));

                    writeSheet.Cells[writeRow, 3] = "'" + elemOut;

                    progressScreen.Increment(range);

                    i = last;

                    writeRow++;
                }
            }

            catch (Exception e)
            {
                //console.Add(e.Message);
                progressScreen.Close();

                return 0;
            }

            return 1;
        }

        private void ExportAreas(object sender, EventArgs e)
        {
            if (Save())
            {
                MessageScreen confirm = new MessageScreen("Please make sure Excel is closed before continuing.");
                if (confirm.ShowDialog() == DialogResult.OK)
                {
                    foreach (Process clsProcess in Process.GetProcessesByName("EXCEL"))
                    {
                        clsProcess.Kill();
                    }

                    UpdateElements();

                    this.Cursor = Cursors.WaitCursor;

                    int numTables = 0;
                    string[] tableKeys = { };
                    string[] tableNames = { };
                    int[] import = { };
                    bool[] empty = { };

                    _SapModel.DatabaseTables.GetAllTables(ref numTables, ref tableKeys, ref tableNames, ref import, ref empty);

                    bool found = false;

                    int pierTable = Array.IndexOf(tableNames, "Design Forces - Piers");
                    int spandrelTable = Array.IndexOf(tableNames, "Design Forces - Spandrels");

                    if (!empty[pierTable] && !empty[spandrelTable])
                    {
                        found = true;

                        string[] tableKeyList = { tableKeys[pierTable], tableKeys[spandrelTable] };
                        _SapModel.DatabaseTables.ShowTablesInExcel(ref tableKeyList, 0);

                        Excel.Application app =
                            System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;

                        app.Visible = false;

                        Workbook workbook = app.ActiveWorkbook;

                        Worksheet pierSheet = workbook.Worksheets.Item[1] as Worksheet;
                        Worksheet spandrelSheet = workbook.Worksheets.Item[2] as Worksheet;
                        Worksheet newSheet = workbook.Worksheets.Add();

                        ProgressBar progressScreen = new ProgressBar(pierSheet.UsedRange.Rows.Count + spandrelSheet.UsedRange.Rows.Count - 6);

                        try
                        {

                            newSheet.Name = "Design Forces - ASCE 41";

                            bool level1 = pierSheet.Cells[4, 4].Columns.Find("PO1") != null;

                            Excel.Range header = pierSheet.Range["A1", "J3"];
                            header.Copy();
                            Excel.Range newHeader = newSheet.Range["A1", "J3"];
                            newHeader.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone);

                            newSheet.Columns["D"].Delete();

                            newSheet.Cells[2, 3].Value = "Elements";

                            progressScreen.Show();

                            WriteAreaSheet(pierSheet, newSheet, progressScreen, app, level1);

                            WriteAreaSheet(spandrelSheet, newSheet, progressScreen, app, level1);

                            newSheet.Cells[1, 13] = "Table: Steel Frame Design Summary - ASCE 41 Analysis";

                            Excel.Range titleHeader = newSheet.Range["A1", "A1"].EntireRow;

                            titleHeader.Interior.Color = newSheet.Range["A1", "A1"].Interior.Color;
                            titleHeader.Font.Bold = true;

                            newSheet.Range["A2", "A2"].EntireRow.Interior.Color = newSheet.Range["A2", "A2"].Interior.Color;
                            newSheet.Range["A2", "A2"].EntireRow.Borders.LineStyle = newSheet.Range["A3", "A3"].Borders.LineStyle;
                            newSheet.Range["A3", "A3"].EntireRow.Interior.Color = newSheet.Range["A3", "A3"].Interior.Color;
                            newSheet.Range["A3", "A3"].EntireRow.Borders.LineStyle = newSheet.Range["A2", "A2"].Borders.LineStyle;
                            newSheet.Range["A2", "A2"].EntireRow.Font.Bold = true;

                            for (int j = 0; j < headers.Count(); ++j)
                            {
                                newSheet.Cells[2, 13 + j] = headers[j];
                            }

                            newSheet.Range["A1"].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                            string end = "Q" + newSheet.UsedRange.Rows.Count;

                            newSheet.Range["M2", "Q2"].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            newSheet.Range["A2", end].Columns.AutoFit();

                            //console.Add(progressScreen.Value() + "/" + progressScreen.Maximum() + "\n" + selectedElems.Count + " selected elements");

                            app.Visible = true;

                            progressScreen.Close();
                        }
                        catch (Exception ex)
                        {
                            //console.Add(ex.Message);
                            new MessageScreen("An error occurred. The Excel sheet could not be generated.").Show();
                            progressScreen.Close();
                        }


                        // interop cleanup
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(pierSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(spandrelSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(newSheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                    }

                    if (!found)
                    {
                        new MessageScreen("Could not find Pier or Spandrels Design Forces table - Has analysis been run?").Show();
                    }

                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void CmBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(CmBox.Text);
                this.cm = num;
                this.SetVW();
            }
            catch (Exception) { }
        }

        private void SetVW()
        {
            this.vw[0] = c1c2 * cm * Sa[0];
            vw1Box.Text = vw[0].ToString("0.00");
            this.vw[1] = c1c2 * cm * Sa[1];
            vw2Box.Text = vw[1].ToString("0.00");
            SetVSHL();
        }   

        private void SetB()
        {
            this.b1 = 4 / (5.6 - Math.Log(100 * this.dampingRatio));
            BBox.Text = b1.ToString("0.000");
            SetSa();
        }

        private void FSystemBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            system = FSystemBox.SelectedIndex;

            switch (system)
            {
                case 0:
                    this.ct = 0.035;
                    this.beta = 0.8;
                    break;
                case 1:
                    this.ct = 0.018;
                    this.beta = 0.9;
                    break;
                case 2:
                    this.ct = 0.03;
                    this.beta = 0.75;
                    break;
                default:
                    this.ct = 0.02;
                    this.beta = 0.75;
                    break;
            }
            CtBox.Text = this.ct.ToString("0.000");
            BetaBox.Text = this.beta.ToString("0.00");
            SetB();
            SetT();
        }

        private void Sxs1Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(Sxs1Box.Text);
                this.Sxs[0] = num;
                SetTs();
            }
            catch (Exception) { }
        }

        private void Sxs2Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(Sxs2Box.Text);
                this.Sxs[1] = num;
                SetTs();
            }
            catch (Exception) { }
        }

        private void Sx1Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(Sx1Box.Text);
                this.Sx1[0] = num;
                SetTs();
            }
            catch (Exception) { }
        }

        private void Sx2Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(Sx2Box.Text);
                this.Sx1[1] = num;
                SetTs();
            }
            catch (Exception) { }
        }

        private void SWeightBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(SWeightBox.Text);
                this.sWeight = num;
                SetVSHL();
            }
            catch (Exception) { }
        }


        private void SetTs()
        {
            try
            {
                if (Sxs[0] != 0)
                {
                    Ts[0] = Sx1[0] / Sxs[0];
                    Ts1Box.Text = Ts[0].ToString("0.000");
                }
                if (Sxs[1] != 0)
                {
                    Ts[1] = Sx1[1] / Sxs[1];
                    Ts2Box.Text = Ts[1].ToString("0.000");
                }
            }
            catch
            {

            }
            SetT0();
        }

        private void SetT0()
        {
            try
            {
                T0[0] = 0.2 * Ts[0];
                T01Box.Text = T0[0].ToString("0.000");
                T0[1] = 0.2 * Ts[1];
                T02Box.Text = T0[1].ToString("0.000");
            }
            catch
            {

            }
            SetSa();
            
        }

        private void SetT()
        {
            if(ct != 0)
            {
                period = ct * Math.Pow(hn, beta);
                Period_Box.Text = period.ToString("0.00");
                SetSa();
            }
        }

        private void SetSa()
        {
            if (b1 != 0)
            {
                for (int i = 0; i < 2; ++i)
                {
                    if (this.period < T0[i])
                    {
                        this.Sa[i] = Sxs[i] * ((5 / this.b1 - 2) * (period / Ts[i]) + 0.4);
                    }
                    else if (this.period < Ts[i])
                    {
                        this.Sa[i] = Sxs[i] / b1;
                    }
                    else if (this.period < 1)
                    {
                        this.Sa[i] = Sx1[i] / (b1 * period);
                    }
                    else
                    {
                        this.Sa[i] = Tl[i] * Sx1[i] / (b1 * Math.Pow(period, 2));
                    }
                }
                Sa1Box.Text = Sa[0].ToString("0.00");
                Sa2Box.Text = Sa[1].ToString("0.00");
                SetVW();
            }
        }

        private void SetVSHL()
        {
            this.VSHL[0] = vw[0] * sWeight;
            Vshl1Box.Text = this.VSHL[0].ToString("0.00");
            
            this.VSHL[1] = vw[1] * sWeight;
            Vshl2Box.Text = this.VSHL[1].ToString("0.00");

            Set85();
        }
        
        private void SetScaleFactor()
        {
            this.scaleFactor[0] = 386.4 * c1c2;
            sFactorx1Box.Text = scaleFactor[0].ToString("0.00");
            sFactory1Box.Text = scaleFactor[0].ToString("0.00");
            this.scaleFactor[1] = 386.4 * c1c2;
            sFactorx2Box.Text = scaleFactor[1].ToString("0.00");
            sFactory2Box.Text = scaleFactor[1].ToString("0.00");
            SetFinalScale();
        }

        private void Set85()
        {
            this.LSP85[0] = VSHL[0] * 0.85;
            LSP85xBox1.Text = LSP85[0].ToString("0.00");
            LSP85yBox1.Text = LSP85[0].ToString("0.00");
            this.LSP85[1] = VSHL[1] * 0.85;
            LSP85xBox2.Text = LSP85[1].ToString("0.00");
            LSP85yBox2.Text = LSP85[1].ToString("0.00");
            SetFinalScale();
        }

        private void iLDPx1Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(iLDPx1Box.Text);
                this.iLDP[0,0] = num;
                SetFinalScale();
            }
            catch (Exception) { }
        }

        private void iLDPy1Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(iLDPy1Box.Text);
                this.iLDP[0, 1] = num;
                SetFinalScale();
            }
            catch (Exception) { }
        }

        private void iLDPx2Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(iLDPx2Box.Text);
                this.iLDP[1,0] = num;
                SetFinalScale();
            }
            catch (Exception) { }
        }

        private void iLDPy2Box_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(iLDPy2Box.Text);
                this.iLDP[1, 1] = num;
                SetFinalScale();
            }
            catch (Exception) { }
        }

        private void SetFinalScale()
        {
            for (int i = 0; i < iLDP.Rank; ++i)
            {
                for (int j = 0; j < iLDP.GetLength(i); ++j)
                {

                    if (iLDP[i, j] >= LSP85[i])
                    {
                        this.finalFactor[i, j] = scaleFactor[i];
                    }
                    else
                    {
                        if (iLDP[i, j] != 0)
                        {
                            this.finalFactor[i, j] = scaleFactor[i] * (LSP85[i] / iLDP[i, j]);
                        }
                        else this.finalFactor[i, j] = 0.0;
                    }
                }
            }
            finalScalex1Box.Text = finalFactor[0, 0].ToString("0.00");
            finalScaley1Box.Text = finalFactor[0, 1].ToString("0.00");
            finalScalex2Box.Text = finalFactor[1, 0].ToString("0.00");
            finalScaley2Box.Text = finalFactor[1, 1].ToString("0.00");
        }

        private void dRatioBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double num = double.Parse(dRatioBox.Text);
                this.dampingRatio = num;
                SetB();
            }
            catch (Exception) { }
        }
    }
}
