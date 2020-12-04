using ExcelDataReader;
using ExtractedHMITool.Properties;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Drawing.Drawing2D;

namespace ExtractedHMITool
{
    public partial class mainForm : Form
    {
        private float zoom = 1f;
        private float sizeRectX;
        private float sizeRectY;
        private float posiRectX;
        private float posiRectY;
        float[] savedZoom;
        Color foreColor;
        Color backColor;
        Button button;
        TextBox tb;
         
        private bool excelLoaded = false;
        private bool rectLoaded = false;
        private bool highlight = false;
        private bool isHighlighted = false;

        bool highlightChecked = false;
        bool fileSaved = true;
        bool searchedAllText = false;
        bool searchAllClicked = false;
        bool landscape = false;
        bool searchingTextbox = false;
        bool resetSearchFiles = false;

        private System.Drawing.Point mouseDown;
        private System.Drawing.Point originPicLocation;
        private System.Drawing.RectangleF[] rects;

        private string fpath;
        string richTextPath = "";
        string selectedSheetName = "";
        string filename = "";
        string firstSheetName = "";
        private string[] backColorFrmXl;
        private string[] foreColorFrmXl;
        private string[] typeFrmXl;
        private string[] nameFrmXl;
        private int[] xPositionFrmXl;
        private int[] yPositionFrmXl;
        private int[] xSizeFrmXl;
        private int[] ySizeFrmXl;
        private string[] configFrmXl;
        private string[] animationFrmXl;
        private string[] sheetNames;
        string[] filePaths;
        string[] checkedFiles;

        private int selectedGobjectIndex;
        private int previousGObject = -1;
        private int rowCount;
        private int colCount;
        private int picMovedPositionX;
        private int CurrentZoom = 0;
        private int picMovedPositionY;
        private int penWidth = 1;
        int previousRectIndex = -1;
        int currentSheetIndex = -1;
        int pbOriginalWidth = 0;
        int pbOriginalHeight = 0;
        int sheetCount = 0;
        int mainDisplaySizeX = 0;
        int mainDisplaySizeY = 0;
        int currentSelectedRow = 0;
        int selectedGroupIndex = 0;
        int selectedMainGroupIndex = 0;
        int selectedEndGroupIndex = 0;
        int selectedMainEndGroupIndex = 0;
        int timerIndex = 0;
        int currentPage = 0;
        int selectionStart = 0;
        int selectionStop = 0;
        int selectionStart2 = 0;
        int selectionStop2 = 0;
        int countSearchFiles = 0;
        int chosenPaperSizeHeight = 0;
        int chosenPaperSizeWidth = 0;
        int marginX = 20;
        int marginY = 60;
        int[] checkedIndices;
        int[] savedX;
        int[] savedY;
        int[] savedWidth;
        int[] savedHeight;

        private Size originPicSize;
        private TreeNode root;
        private PictureBox pic;
        private Brush brush;
        private Pen pen;
        private Pen penLine;
        private FileStream stream;
        private DataSet result;
        private DataTable dt;
        private IExcelDataReader excelReader;
        private Font font;
        private Point currentPicBoxLocation;
        detailForm form = null;
        Point detailFormLocation;
        printForm pForm;
        FileInfo fileInfo;
        ExcelPackage package;
        PrintPreviewDialog preview;
        RectangleF buttonRect;
        Brush linGrBrushX = null;
        Brush linGrBrushY = null;

        public mainForm()
        {
            InitializeComponent();
        }

        private void pictureBoxMain_Paint(object sender, PaintEventArgs e)
        {
            linGrBrushX = null;
            linGrBrushY = null;

            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // uploads images and draw shapes only if excel is loaded
            if (excelLoaded)
            {
                if (backColorFrmXl[0].Contains("#"))
                {
                    pictureBoxMain.BackColor = ColorTranslator.FromHtml(backColorFrmXl[0]);
                }
                else
                {
                    pictureBoxMain.BackColor = Color.FromName(backColorFrmXl[0]);

                }

                //draw rectangle for all objects to be clickable
                rects = new System.Drawing.RectangleF[rowCount];
                string[] temp = null;
                PointF[] points = null;

                for (int i = 0; i < rowCount; i++)
                {
                    //determine if the color is in string format or HEX format
                    if (foreColorFrmXl[i].Contains("#"))
                    {
                        pen = new Pen(ColorTranslator.FromHtml(foreColorFrmXl[i]));
                        penLine = new Pen(ColorTranslator.FromHtml(foreColorFrmXl[i]), penWidth);
                        foreColor = ColorTranslator.FromHtml(foreColorFrmXl[i]);
                    }
                    else
                    {
                        pen = new Pen(Color.FromName(foreColorFrmXl[i]));
                        penLine = new Pen(Color.FromName(foreColorFrmXl[i]), penWidth);
                        foreColor = Color.FromName(foreColorFrmXl[i]);
                    }

                    if (backColorFrmXl[i].Contains("#"))
                    {
                        brush = new SolidBrush(ColorTranslator.FromHtml(backColorFrmXl[i]));
                        backColor = ColorTranslator.FromHtml(backColorFrmXl[i]);

                    }
                    else
                    {
                        brush = new SolidBrush(Color.FromName(backColorFrmXl[i]));
                        backColor = Color.FromName(backColorFrmXl[i]);

                    }

                    // substring astericks from object type column
                    string objectType = "";

                    if (typeFrmXl[i].Contains("circle"))
                    {
                        objectType = "circle";
                    }
                    else if (typeFrmXl[i].Contains("line"))
                    {
                        objectType = "line";
                    }
                    else if (typeFrmXl[i].Contains("rect"))
                    {
                        objectType = "rectangle";
                    }
                    else if (typeFrmXl[i].Contains("image"))
                    {
                        objectType = "image";
                    }
                    else if (typeFrmXl[i].Contains("string"))
                    {
                        objectType = "string";
                    }
                    else if (typeFrmXl[i].Contains("polygon"))
                    {
                        objectType = "polygon";
                    }
                    else if (typeFrmXl[i].Contains("input"))
                    {
                        objectType = "input";
                    }
                    else if (typeFrmXl[i].Contains("textbox"))
                    {
                        objectType = "textbox";
                    }
                    else
                    {
                        objectType = "group";
                    }




                    if (objectType.Equals("line") || objectType.Equals("polygon"))
                    {
                        string[] strArray = (configFrmXl[i]).Split(new string[] { "(||)" }, StringSplitOptions.None);
                        string pointsStr = strArray[0];
                        string lineWidthStr = strArray[1];
                        int lineWidth = Convert.ToInt32(lineWidthStr);
                        penLine.Width = lineWidth;
                        temp = pointsStr.Split(' ');
                        points = new PointF[temp.Length / 2];
                        int pointsCounter = 0;

                        for (int index = 0; index < temp.Length; index += 2)
                        {
                            points[pointsCounter] = new PointF(float.Parse(temp[index]) * zoom, float.Parse(temp[index + 1]) * zoom);
                            pointsCounter++;
                            penLine.Width *= zoom;
                            if (pointsCounter == temp.Length / 2)
                            {
                                break;
                            }
                        }

                       
                    }
                    rects[i] = new System.Drawing.RectangleF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom,
                            xSizeFrmXl[i] * zoom, ySizeFrmXl[i] * zoom);


                    //determine the types(shapes) of the object
                    switch (objectType)
                    {
                        case "rectangle":
                            if(configFrmXl[i].Contains("gradient") && configFrmXl[i].Contains("rotation"))
                            {
                                if(xSizeFrmXl[i]<ySizeFrmXl[i])
                                {                                                                       
                                    if(xSizeFrmXl[i] != 0)
                                    {
                                        using (linGrBrushX = new LinearGradientBrush(new PointF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom),
                                        new PointF((xPositionFrmXl[i] + xSizeFrmXl[i]) * zoom, (yPositionFrmXl[i]) * zoom), Color.White, backColor))
                                        {
                                            double degree = Convert.ToDouble(configFrmXl[i].Replace("gradient(||)rotation(||)", ""));
                                            RotateRectangle(g, rects[i], (float)degree, linGrBrushX);
                                        }
                                    }                                
                                }
                                else
                                {
                                    using ( linGrBrushY = new LinearGradientBrush(new PointF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom),
                                        new PointF((xPositionFrmXl[i]) * zoom, (yPositionFrmXl[i] + ySizeFrmXl[i]) * zoom), Color.White, backColor))
                                    {
                                        double degree = Convert.ToDouble(configFrmXl[i].Replace("gradient(||)rotation(||)", ""));
                                        RotateRectangle(g, rects[i], (float)degree, linGrBrushY);
                                    }
                                }

                            }
                            else if(configFrmXl[i].Contains("gradient"))
                            {
                                if(xSizeFrmXl[i] < ySizeFrmXl[i])
                                {
                                    if (xSizeFrmXl[i] != 0)
                                    {
                                        using (linGrBrushX = new LinearGradientBrush(new PointF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom),
                                        new PointF((xPositionFrmXl[i] + xSizeFrmXl[i]) * zoom, (yPositionFrmXl[i]) * zoom), Color.White, backColor))
                                        {
                                            
                                            g.FillRectangle(linGrBrushX, rects[i]);
                                        }
                                    }
                                }
                                else
                                {
                                    using (linGrBrushY = new LinearGradientBrush(new PointF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom),
                                        new PointF((xPositionFrmXl[i]) * zoom, (yPositionFrmXl[i] + ySizeFrmXl[i]) * zoom), Color.White, backColor))
                                    {
                                        
                                        g.FillRectangle(linGrBrushY, rects[i]);
                                    }
                                }
                            }
                            else if (configFrmXl[i].Contains("rotation"))
                            {

                                double degree = Convert.ToDouble(configFrmXl[i].Replace("rotation(||)", ""));
                                RotateRectangle(g, rects[i], (float)degree, new SolidBrush(backColor));

                            }
                            else
                            {
                                g.DrawRectangle(penLine, xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom, xSizeFrmXl[i] * zoom, ySizeFrmXl[i] * zoom);
                                g.FillRectangle(brush, rects[i]);
                            }
                            
                            break;
                        case "circle":
                            if (configFrmXl[i].Contains("gradient"))
                            {
                                if (xSizeFrmXl[i] < ySizeFrmXl[i])
                                {
                                    if (xSizeFrmXl[i] != 0)
                                    {
                                        using (linGrBrushX = new LinearGradientBrush(new PointF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom),
                                        new PointF((xPositionFrmXl[i] + xSizeFrmXl[i]) * zoom, (yPositionFrmXl[i]) * zoom), Color.White, backColor))
                                        {
                                            g.FillEllipse(linGrBrushX, rects[i]);
                                        }
                                    }
                                }
                                else
                                {
                                    using (linGrBrushY = new LinearGradientBrush(new PointF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom),
                                        new PointF((xPositionFrmXl[i]) * zoom, (yPositionFrmXl[i] + ySizeFrmXl[i]) * zoom), Color.White, backColor))
                                    {
                                        g.FillEllipse(linGrBrushY, rects[i]);

                                    }
                                }

                            }
                            else
                            {
                                g.DrawEllipse(penLine, rects[i]);
                                g.FillEllipse(brush, rects[i]);
                            }
                               
                            break;
                        case "image":
                            loadImgByPath(e, configFrmXl[i].ToString(), i);
                            break;
                        case "string":
                            string[] strArray = (configFrmXl[i]).Split(new string[] { "(||)" }, StringSplitOptions.None);
                            string fontType = strArray[0];
                            string fontSize = strArray[1];
                            string fontStyle = strArray[2];
                            string text = strArray[3];
                            if(fontStyle.CaseInsensitiveContains("bold"))
                            {
                                font = new Font(fontType, Convert.ToInt32(fontSize) * zoom, FontStyle.Bold);

                            }
                            else
                            {
                                font = new Font(fontType, Convert.ToInt32(fontSize) * zoom, FontStyle.Regular);

                            }
                            //g.DrawString(text, font, brush, xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom);
                            g.DrawString(text, font, brush, new RectangleF(xPositionFrmXl[i]*zoom,yPositionFrmXl[i]*zoom,(xSizeFrmXl[i] + 15)*zoom,ySizeFrmXl[i]*zoom));
                            break;
                        case "line":
                            penLine.EndCap = LineCap.ArrowAnchor;
                            g.DrawLines(penLine, points);
                            penLine.EndCap = LineCap.NoAnchor;

                            //g.FillPolygon(brush, points);
                            break;
                        case "polygon":

                            g.DrawPolygon(penLine, points);
                            g.FillPolygon(brush, points);
                            break;

                        case "input":
                            string[] strArrayB = (configFrmXl[i]).Split(new string[] { "(||)" }, StringSplitOptions.None);
                            string fontTypeB = strArrayB[0];
                            string fontSizeB = strArrayB[1];
                            string fontStyleB = strArrayB[2];
                            string textB = strArrayB[3];
                            if (fontStyleB.CaseInsensitiveContains("bold"))
                            {
                                font = new Font(fontTypeB, Convert.ToInt32(fontSizeB) * zoom, FontStyle.Bold);

                            }
                            else
                            {
                                font = new Font(fontTypeB, Convert.ToInt32(fontSizeB) * zoom, FontStyle.Regular);

                            }
                            buttonRect = new RectangleF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom, (xSizeFrmXl[i]) * zoom, ySizeFrmXl[i] * zoom);
                            
                            g.DrawRectangle(new Pen(ColorTranslator.FromHtml("#f9f9f9"),3), xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom, xSizeFrmXl[i] * zoom, ySizeFrmXl[i] * zoom);
                            g.FillRectangle(brush, buttonRect);
                            StringFormat sf = new StringFormat();
                            sf.LineAlignment = StringAlignment.Center;
                            sf.Alignment = StringAlignment.Center;
                            //g.DrawString(text, font, brush, xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom);
                            g.DrawString(textB, font, new SolidBrush(foreColor), buttonRect,sf);
                            break;

                        case "textbox":

                            string[] strArrayT = (configFrmXl[i]).Split(new string[] { "(||)" }, StringSplitOptions.None);
                            string fontTypeT = strArrayT[0];
                            string fontSizeT = strArrayT[1];
                            string fontStyleT = strArrayT[2];
                            string textT = strArrayT[3];
                            if (fontStyleT.CaseInsensitiveContains("bold"))
                            {
                                font = new Font(fontStyleT, Convert.ToInt32(fontSizeT) * zoom, FontStyle.Bold);

                            }
                            else
                            {
                                font = new Font(fontStyleT, Convert.ToInt32(fontSizeT) * zoom, FontStyle.Regular);

                            }

                            buttonRect = new RectangleF(xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom, (xSizeFrmXl[i]) * zoom, ySizeFrmXl[i] * zoom);

                            g.DrawRectangle(new Pen(Color.Black, 1), xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom, xSizeFrmXl[i] * zoom, ySizeFrmXl[i] * zoom);
                            g.FillRectangle(brush, buttonRect);
                            StringFormat sf2 = new StringFormat();
                            sf2.LineAlignment = StringAlignment.Center;
                            sf2.Alignment = StringAlignment.Center;
                            //g.DrawString(text, font, brush, xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom);
                            g.DrawString(textT, font, new SolidBrush(foreColor), buttonRect, sf2);
                            break;

                    }
                }
                rectLoaded = true;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private void loadImgByPath(PaintEventArgs e, string imageFile, int i)
        {
            try
            {
                if (!imageFile.Equals(""))
                {


                    string str = fpath.Replace(filename, "");
                    Image img = Image.FromFile(imageFile);
                    Graphics g = e.Graphics;
                    g.DrawImage(img, xPositionFrmXl[i] * zoom, yPositionFrmXl[i] * zoom, xSizeFrmXl[i] * zoom, ySizeFrmXl[i] * zoom);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void excelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                //if not saved
                if (!fileSaved)
                {
                    DialogResult dialogResult = MessageBox.Show("Do you want to save the current file before opening a new file?", "Save", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.Yes)
                    {
                        package.Save();
                        statusLbl.Text = "File is saved";
                        var t = new System.Windows.Forms.Timer();
                        t.Interval = 3000; // it will Tick in 3 seconds
                        t.Tick += (s, p) =>
                        {
                            statusLbl.Text = "Ready";
                            t.Stop();
                        };
                        t.Start();

                        fileSaved = true;

                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        //do not save

                    }
                }

                //upload excel by open file dialog
                fpath = "";
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Title = "Open Excel File";
                fdlg.InitialDirectory = "";
                fdlg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    fpath = fdlg.FileName;
                    filename = fdlg.SafeFileName;
                }

                if (fpath.Equals(""))
                {
                    //if User closes the file dialog
                    excelLoaded = false;
                    MessageBox.Show("Alert: Open Excel file first!");
                    explorer.Nodes.Clear();
                    pictureBoxMain.Visible = false;
                    ExcelDataGrid.DataSource = null;
                    fileSaved = true;

                }
                else
                {
                    excelLoaded = false;
                    pictureBoxMain.Visible = false;
                    ExcelDataGrid.Visible = false;
                    ExcelDataGrid.DataSource = null;
                    fileSaved = true;

                    ImageList myImageList = new ImageList();
                    myImageList.Images.Add(Resources.Excel_Icon);
                    myImageList.Images.Add(Resources.sheet);

                    explorer.ImageList = myImageList;

                    explorer.Nodes.Clear();

                    //set the text of the root node
                    root = new TreeNode(fpath, 0, 0);
                    explorer.Nodes.Add(root);

                    //clear rich text box
                    animationTextBox.Clear();

                    //perform reset position,zoom
                    resetBtn.PerformClick();

                    using (stream = new FileStream(fpath, FileMode.Open, FileAccess.Read))
                    {
                        // read from excel file
                        excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                        // add treenodes for excel sheets
                        TreeNode[] sheets = new TreeNode[excelReader.ResultsCount];
                        sheetNames = new string[excelReader.ResultsCount];
                        int count = 0;
                        //popluate the tree with excel sheet name info
                        for (int i = 0; i < excelReader.ResultsCount; i++)
                        {

                            if (i > 0)
                            {
                                //if more sheets exist
                                if (excelReader.NextResult())
                                {
                                    sheets[i] = new TreeNode(excelReader.Name, 1, 1);
                                    if (excelReader.Name.Contains("Grp"))
                                    {
                                        root.Nodes.Add(sheets[i]);
                                        if (count == 0)
                                        {
                                            firstSheetName = excelReader.Name;
                                            count++;
                                        }

                                    }
                                    sheetNames[i] = excelReader.Name;

                                }
                            }
                            else
                            {
                                sheets[i] = new TreeNode(excelReader.Name, 1, 1);
                                if (excelReader.Name.Contains("Grp"))
                                {
                                    root.Nodes.Add(sheets[i]);
                                    firstSheetName = excelReader.Name;

                                }
                                sheetNames[i] = excelReader.Name;
                            }
                        }
                        result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            // Gets or sets a value indicating whether to set the DataColumn.DataType
                            // property in a second pass.
                            UseColumnDataType = true,

                            // Gets or sets a callback to obtain configuration options for a DataTable.
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                // Gets or sets a value indicating whether to use a row from the
                                // data as column names.
                                UseHeaderRow = true,
                            }
                        });

                    }
                    ExcelDataGrid.DataSource = null;
                    explorer.ExpandAll();

                    sheetCount = sheetNames.Length;
                    savedX = new int[sheetCount];
                    savedY = new int[sheetCount];
                    savedWidth = new int[sheetCount];
                    savedHeight = new int[sheetCount];
                    savedZoom = new float[sheetCount];

                    fileInfo = new FileInfo(fpath);
                    package = new ExcelPackage(fileInfo);
                    saveScreenshots();
                    filePaths = Directory.GetFiles(fpath.Replace(filename, "") + "\\screenshots\\", "*.png", SearchOption.TopDirectoryOnly);
                    highlightCheckBox.Enabled = true;
                    XlSheet_loadOnClick(firstSheetName);
                }
            }
            catch (IOException error)
            {
                System.Windows.Forms.MessageBox.Show(error.ToString());
            }

            // need to close the stream at one point (when opening another excelfile)
            Application.UseWaitCursor = false;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void explorer_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                //clear any highlighted objects
                pictureBoxMain.Controls.Clear();
                selectedSheetName = e.Node.Text;

                if (selectedSheetName != fpath)
                {
                    XlSheet_loadOnClick(selectedSheetName);
                }

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void pictureBoxMain_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                CursorLocationLabel.Text = "Current Cursor Position: " + e.Location;

                if (previousRectIndex == -1)
                {
                    pictureBoxMain.Cursor = Cursors.Default;

                    if (rectLoaded)
                    {

                        for (int j = 0; j < rowCount; j++)
                        {
                            if (rects[j].Contains(e.X, e.Y))
                            {
                                pictureBoxMain.Cursor = Cursors.Hand;
                                previousRectIndex = j;

                            }
                        }

                    }
                }
                else
                {
                    if (!rects[previousRectIndex].Contains(e.X, e.Y))
                    {
                        previousRectIndex = -1;
                    }

                }

                if (e.Button == MouseButtons.Left)
                {
                    Point mousePosNow = e.Location;

                    int deltaX = mousePosNow.X - mouseDown.X;
                    int deltaY = mousePosNow.Y - mouseDown.Y;

                    picMovedPositionX = pictureBoxMain.Location.X + deltaX;
                    picMovedPositionY = pictureBoxMain.Location.Y + deltaY;

                    pictureBoxMain.Location = new Point(picMovedPositionX, picMovedPositionY);
                    currentPicBoxLocation = new Point(picMovedPositionX, picMovedPositionY);
                    //cleanup  
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void pictureBoxMain_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                mouseDown = e.Location;
            }
        }

        //Enable mousewheel to zoom in and zoom out
        protected override void OnMouseWheel(MouseEventArgs e)
        {
            try
            {
                if ((ModifierKeys & Keys.Control) == Keys.Control && excelLoaded)
                {
                    if (e.Delta > 0)
                    {

                        if (zoom > 1.7)

                        {
                            System.Windows.Forms.MessageBox.Show("You have exceed the maximum number of zoom in.");

                        }
                        else
                        {
                            //resize the picturebox1
                            zoom += 0.1f;
                            pictureBoxMain.Size = new Size((int)(pbOriginalWidth * zoom), (int)(pbOriginalHeight * zoom));
                            pictureBoxMain.Location = new System.Drawing.Point(picMovedPositionX, picMovedPositionY);
                            pictureBoxMain.Invalidate();
                            //clear the highlight
                            pictureBoxMain.Controls.Clear();
                        }
                    }
                    else if (e.Delta < 0)
                    {

                        if (zoom < 0.3)
                        {
                            System.Windows.Forms.MessageBox.Show("You have exceed the maximum number of zoom out.");
                        }
                        else
                        {
                            zoom -= 0.1f;

                            pictureBoxMain.Size = new Size((int)(pbOriginalWidth * zoom), (int)(pbOriginalHeight * zoom));
                            pictureBoxMain.Location = new System.Drawing.Point(picMovedPositionX, picMovedPositionY);
                            //redraw the picturebox
                            pictureBoxMain.Invalidate();
                            //clear the highlight
                            pictureBoxMain.Controls.Clear();
                        }
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void pictureBoxMain_MouseClick(object sender, MouseEventArgs e)
        {
            int overlapCount = 0;
            int index = 0;
            try
            {
                pictureBoxMain.Enabled = true;

                if (isHighlighted)
                {
                    //clear the highlight
                    pictureBoxMain.Controls.Clear();
                    isHighlighted = false;
                }
                if (e.Button == MouseButtons.Left)
                {
                    if (excelLoaded)
                    {
                        pictureBoxMain.Controls.Clear();


                        for (int j = 0; j < rowCount; j++)
                        {
                            ExcelDataGrid.Rows[j].Selected = false;

                            string objectType = null;

                            if (ExcelDataGrid.Rows[j].Visible)
                            {
                                if (typeFrmXl[j].Contains("*"))
                                {
                                    objectType = typeFrmXl[j].Substring(1);

                                    if (typeFrmXl[j].Contains("endgroup"))
                                    {
                                        objectType = typeFrmXl[j].Replace("endgroup", "");
                                    }

                                }

                                else
                                {
                                    if (typeFrmXl[j].Contains("endgroup"))
                                    {
                                        objectType = typeFrmXl[j].Replace("endgroup", "");
                                    }
                                    else
                                    {
                                        objectType = typeFrmXl[j];

                                    }

                                }

                                if (rects[j].Contains(e.X, e.Y))
                                {
                                    if (objectType.Equals("circle") || objectType.Equals("input") || objectType.Equals("textbox") ||objectType.Equals("group") || objectType.Equals("polygon") || objectType.Equals("image") || objectType.Equals("line") || objectType.Equals("string"))
                                    {
                                        if (overlapCount > 0)
                                        {
                                            pictureBoxMain.Controls.Clear();

                                            index = j;

                                        }
                                        previousRectIndex = j;


                                        pictureBoxMain.Controls.Clear();


                                        highlight_OnClick(j);

                                        isHighlighted = true;
                                        overlapCount++;

                                    }
                                    changeDetailForm(j);
                                    setDetailPanel();

                                    //move to the selected row
                                    ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[j].Cells[0];
                                    ExcelDataGrid.Rows[j].Selected = true;
                                }
                                else
                                {
                                    pictureBoxMain.Invalidate();
                                }

                            }

                        }
                        if (overlapCount > 1)
                        {
                            if (rects[previousRectIndex].Contains(e.X, e.Y))
                            {
                                pictureBoxMain.Controls.Clear();

                                highlight_OnClick(index);
                            }
                        }
                    }
                }

                //grid highlight
                gridHighlight();

                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception error)
            {
                System.Windows.Forms.MessageBox.Show(error.ToString());
            }

        }

        private void ExcelDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // System.Windows.Forms.MessageBox.Show(ExcelDataGrid.CurrentCell.Value.ToString());

            try
            {
                ExcelDataGrid.Focus();
                ExcelDataGrid.CurrentRow.Selected = true;

                if (excelLoaded)
                {
                    //move the object out of sight to the center position
                    reCenterPosition();

                    //highlight grid
                    gridHighlight();

                    //checkbox control
                    if (e.ColumnIndex == 0)
                    {

                        if (!typeFrmXl[ExcelDataGrid.CurrentRow.Index].Equals("group"))
                        {
                            ExcelDataGrid.BeginEdit(true);
                            fileSaved = false;
                        }
                        else if (!ExcelDataGrid[0, selectedMainGroupIndex].Value.Equals("true"))
                        {
                            for (int i = selectedMainGroupIndex; i <= selectedMainEndGroupIndex; i++)
                            {
                                ExcelDataGrid[0, i].Value = false;

                            }
                        }
                    }

                    selectedGobjectIndex = ExcelDataGrid.CurrentRow.Index;
                    //show rich textbox when clicked on a cell
                    changeDetailForm(selectedGobjectIndex);
                    setDetailPanel();

                    posiRectX = rects[selectedGobjectIndex].Location.X;
                    posiRectY = rects[selectedGobjectIndex].Location.Y;
                    sizeRectX = rects[selectedGobjectIndex].Size.Width;
                    sizeRectY = rects[selectedGobjectIndex].Size.Height;

                    //highlight
                    if (selectedGobjectIndex != previousGObject && !highlight)
                    {

                        highlight_OnClick(selectedGobjectIndex);
                        previousGObject = ExcelDataGrid.CurrentRow.Index;
                        highlight = true;

                    }
                    else if (selectedGobjectIndex == previousGObject && highlight)
                    {
                        highlight_OnClick(selectedGobjectIndex);

                        previousGObject = ExcelDataGrid.CurrentRow.Index;
                        highlight = true;
                    }
                    else if (selectedGobjectIndex != previousGObject && highlight)
                    {
                        //unhighlight the previous node
                        pictureBoxMain.Controls.Clear();
                        //highlight the current node
                        highlight_OnClick(selectedGobjectIndex);

                    }
                    else
                    {
                        pictureBoxMain.Controls.Clear();
                    }

                }
            }
            catch (Exception error)
            {
                System.Windows.Forms.MessageBox.Show(error.ToString());
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }
        private void highlight_OnClick(int counter)
        {
            pictureBoxMain.Invalidate();
            //move to the selected row
            ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[counter].Cells[0];
            ExcelDataGrid.Rows[counter].Selected = true;
            pic.BackColor = Color.Transparent;
            pic.SizeMode = PictureBoxSizeMode.AutoSize;

            //highlight the one clicked
            posiRectX = rects[counter].Location.X;
            posiRectY = rects[counter].Location.Y;
            sizeRectX = rects[counter].Size.Width;
            sizeRectY = rects[counter].Size.Height;

            // if there is no size do not highlight
            if (!typeFrmXl[counter].Equals("display"))
            {
                if (xSizeFrmXl[counter] != 0 || ySizeFrmXl[counter] != 0)
                {
                    Bitmap bm = new Bitmap((int)(sizeRectX + 20 + 2 * CurrentZoom), (int)(sizeRectY + 20 + 2 * CurrentZoom));

                    Pen highlightPen = new Pen(Color.Black, 2);
                    highlightPen.DashPattern = new float[] { 3, 1 };

                    using (Graphics gr = Graphics.FromImage(bm))
                    {
                        Rectangle rectangle = new Rectangle(0, 0, (int)(sizeRectX + 20 + 2 * CurrentZoom), (int)(sizeRectY + 20 + 2 * CurrentZoom));
                        gr.DrawRectangle(highlightPen, rectangle);
                    }
                    pic.Image = bm;

                    //pic.Size = new Size((int)(sizeRectX + 20 + 2 * CurrentZoom), (int)(sizeRectY + 20 + 2 * CurrentZoom));
                    pic.Location = new System.Drawing.Point((int)posiRectX - 10 - 1 * CurrentZoom, (int)posiRectY - 10 - 1 * CurrentZoom);
                }
            }

            pictureBoxMain.Controls.Add(pic);
            isHighlighted = true;

        }
        private void multipleHighlight_OnClick(int counter)
        {
            //move to the selected row
            ExcelDataGrid.Rows[counter].Selected = true;

            //highlight the one clicked
            posiRectX = rects[counter].Location.X;
            posiRectY = rects[counter].Location.Y;
            sizeRectX = rects[counter].Size.Width;
            sizeRectY = rects[counter].Size.Height;

            Bitmap bm = new Bitmap((int)(sizeRectX + 20 + 2 * CurrentZoom), (int)(sizeRectY + 20 + 2 * CurrentZoom));

            Pen highlightPen = new Pen(Color.Black, 2);
            highlightPen.DashPattern = new float[] { 3, 1 };

            using (Graphics gr = Graphics.FromImage(bm))
            {
                Rectangle rectangle = new Rectangle(0, 0, (int)(sizeRectX + 20 + 2 * CurrentZoom), (int)(sizeRectY + 20 + 2 * CurrentZoom));
                gr.DrawRectangle(highlightPen, rectangle);
            }

            Point point = new System.Drawing.Point((int)posiRectX - 10 - 1 * CurrentZoom, (int)posiRectY - 10 - 1 * CurrentZoom);
            Graphics g = pictureBoxMain.CreateGraphics();

            if (!typeFrmXl[counter].Equals("display"))
            {
                if (!xSizeFrmXl[counter].Equals("0") || !ySizeFrmXl[counter].Equals("0"))
                    g.DrawImage(bm, point);
            }
        }

        private void XlSheet_loadOnClick(string sheetName)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                pictureBoxMain.Visible = false;
                pictureBoxMain.Enabled = false;
                ExcelDataGrid.Visible = false;

                int index = 0;
                int x = pictureBoxMain.Location.X;
                int y = pictureBoxMain.Location.Y;
                int width = pictureBoxMain.Size.Width;
                int height = pictureBoxMain.Size.Height;
                highlightCheckBox.Checked = false;

                //clear rich textbox
                animationTextBox.Clear();

                if (form.Visible)
                {
                    form.path_RTF = "";
                    int posiX = this.Location.X;
                    int posiY = this.Location.Y;
                    form.Location = new Point(posiX, posiY + 80);
                }

                if (currentSheetIndex == -1)
                {
                    pictureBoxMain.Location = new System.Drawing.Point(0, 0);
                    zoom = 1;
                }
                // save the position,size,zoom of the previous sheet
                else
                {
                    savedX[currentSheetIndex] = x;
                    savedY[currentSheetIndex] = y;
                    savedWidth[currentSheetIndex] = width;
                    savedHeight[currentSheetIndex] = height;
                    savedZoom[currentSheetIndex] = zoom;

                }

                for (int i = 0; i < sheetCount; i++)
                {
                    if (sheetNames[i].Equals(sheetName))
                    {
                        index = i;
                    }
                }

                currentSheetIndex = index;

                // pick the table using the name of the sheet
                dt = result.Tables[index];

                //remove first row to make it easier to convert
                rowCount = dt.Rows.Count;
                colCount = dt.Columns.Count;

                ExcelDataGrid.Visible = true;

                //Set the data source to gridview
                ExcelDataGrid.DataSource = dt;

                typeFrmXl = new string[rowCount];
                foreColorFrmXl = new string[rowCount];
                backColorFrmXl = new string[rowCount];
                xPositionFrmXl = new int[rowCount];
                yPositionFrmXl = new int[rowCount];
                xSizeFrmXl = new int[rowCount];
                ySizeFrmXl = new int[rowCount];
                nameFrmXl = new string[rowCount];
                configFrmXl = new string[rowCount];
                animationFrmXl = new string[rowCount];

                for (int i = 0; i < rowCount; i++)
                {
                    typeFrmXl[i] = dt.Rows[i][1].ToString();
                    nameFrmXl[i] = dt.Rows[i][2].ToString();
                    foreColorFrmXl[i] = dt.Rows[i][3].ToString();
                    backColorFrmXl[i] = dt.Rows[i][4].ToString();
                    xPositionFrmXl[i] = Convert.ToInt32(dt.Rows[i][5]);
                    yPositionFrmXl[i] = Convert.ToInt32(dt.Rows[i][6]);
                    xSizeFrmXl[i] = Convert.ToInt32(dt.Rows[i][7]);
                    ySizeFrmXl[i] = Convert.ToInt32(dt.Rows[i][8]);
                    configFrmXl[i] = dt.Rows[i][9].ToString();
                    animationFrmXl[i] = dt.Rows[i][10].ToString();
                    ExcelDataGrid[0, i] = new DataGridViewCheckBoxCell();
                }

                mainDisplaySizeX = Convert.ToInt32(xSizeFrmXl[0]);
                mainDisplaySizeY = Convert.ToInt32(ySizeFrmXl[0]);

                excelLoaded = true;
                rectLoaded = false;
                pictureBoxMain.Visible = true;

                foreach (DataGridViewColumn dc in ExcelDataGrid.Columns)
                {
                    if (dc.Index.Equals(0))
                    {
                        dc.ReadOnly = false;
                    }
                    else
                    {
                        dc.ReadOnly = true;
                    }
                }

                ExcelDataGrid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                ExcelDataGrid.Columns["Object"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ExcelDataGrid.Columns["Name"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ExcelDataGrid.Columns["Check"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ExcelDataGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ExcelDataGrid.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                ExcelDataGrid.Columns[10].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;




                // clicked on the sheet for the first time
                if (savedWidth[currentSheetIndex] == 0)
                {
                    pictureBoxMain.Size = new Size(mainDisplaySizeX, mainDisplaySizeY);
                    pictureBoxMain.Location = new System.Drawing.Point(0, 0);
                    zoom = 1;
                }
                // load the saved position,size,zoom
                else
                {
                    pictureBoxMain.Size = new Size(savedWidth[currentSheetIndex], savedHeight[currentSheetIndex]);
                    pictureBoxMain.Location = new System.Drawing.Point(savedX[currentSheetIndex], savedY[currentSheetIndex]);
                    zoom = savedZoom[currentSheetIndex];
                }

                pictureBoxMain.Invalidate();
                // Clear initial selection
                ExcelDataGrid[0, 0].Selected = false;

                if (!backgroundWorker1.IsBusy)
                    backgroundWorker1.RunWorkerAsync();

                pbOriginalWidth = Convert.ToInt32(xSizeFrmXl[0]);
                pbOriginalHeight = Convert.ToInt32(ySizeFrmXl[0]);

            }
            catch (IOException error)
            {
                System.Windows.Forms.MessageBox.Show(error.ToString());
            }
        }

        private void mainForm_Load(object sender, EventArgs e)
        {
            try
            {
                //set the image of refresh button
                originPicSize = pictureBoxMain.Size;
                originPicLocation = pictureBoxMain.Location;

                //set dataGrid Header color
                ExcelDataGrid.EnableHeadersVisualStyles = false;
                ExcelDataGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.Gray;

                //initalize form and picturebox
                form = new detailForm();
                pictureBoxMain.Visible = false;
                pic = new PictureBox();
                tb = new TextBox();
                button = new Button();

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void plusDetailBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    //clear highlights
                    animationTextBox.SelectAll();
                    animationTextBox.SelectionBackColor = Color.White;
                    animationTextBox.DeselectAll();

                    if (form.IsDisposed)
                    {
                        form = new detailForm();
                        int x = this.Location.X;
                        int y = this.Location.Y;
                        form.Location = new Point(x, y + 80);
                        
                    }
                    form.Owner = this;
                    changeDetailForm(ExcelDataGrid.CurrentRow.Index);                    
                    form.Show();
                    this.CheckKeyword("property", Color.Purple, 0);
                    this.CheckKeyword("name", Color.Green, 0);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void highlight_grid(int mainGroupIndex, int groupIndex, int mainEndGroupIndex, int endGroupIndex)
        {

            for (int i = mainGroupIndex; i <= mainEndGroupIndex; i++)
            {
                if (i >= groupIndex)
                {
                    ExcelDataGrid.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    if (i > endGroupIndex)
                    {
                        ExcelDataGrid.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                    }
                }
                else
                {
                    ExcelDataGrid.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                }
            }
        }

        private void highlightCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[ExcelDataGrid.DataSource];
                    currencyManager1.SuspendBinding();

                    if (highlightCheckBox.Checked)
                    {
                        highlightChecked = true;
                        Cursor.Current = Cursors.WaitCursor;

                        for (int i = 0; i < rowCount; i++)
                        {
                            if (animationFrmXl[i].Equals(""))
                            {
                                ExcelDataGrid.Rows[i].Visible = false;
                            }
                        }
                    }
                    else
                    {
                        highlightChecked = false;
                        Cursor.Current = Cursors.WaitCursor;

                        for (int i = 0; i < rowCount; i++)
                        {
                            if (animationFrmXl[i].Equals(""))
                            {
                                ExcelDataGrid.Rows[i].Visible = true;
                            }
                        }
                    }
                    currencyManager1.ResumeBinding();
                    Cursor.Current = Cursors.Default;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void ZoomInBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    if (zoom > 1.7)
                    {
                        System.Windows.Forms.MessageBox.Show("You have exceed the maximum number of zoom in.");
                    }
                    else
                    {
                        //resize the picturebox1
                        zoom += 0.1f;
                        pictureBoxMain.Size = new Size((int)(pbOriginalWidth * zoom), (int)(pbOriginalHeight * zoom));
                        pictureBoxMain.Location = new System.Drawing.Point(picMovedPositionX, picMovedPositionY);
                        pictureBoxMain.Invalidate();

                        //clear the highlight
                        pictureBoxMain.Controls.Clear();
                    }
                }
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception error)
            {
                System.Windows.Forms.MessageBox.Show(error.ToString());
            }
        }

        private void ZoomOutBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    if (zoom < 0.3)
                    {
                        System.Windows.Forms.MessageBox.Show("You have exceed the maximum number of zoom out.");
                    }
                    else
                    {
                        zoom -= 0.1f;

                        pictureBoxMain.Size = new Size((int)(pbOriginalWidth * zoom), (int)(pbOriginalHeight * zoom));
                        pictureBoxMain.Location = new System.Drawing.Point(picMovedPositionX, picMovedPositionY);

                        //redraw the picturebox
                        pictureBoxMain.Invalidate();

                        //clear the highlight
                        pictureBoxMain.Controls.Clear();
                    }
                }
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception error)
            {
                System.Windows.Forms.MessageBox.Show(error.ToString());
            }
        }

        private void resetBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    pictureBoxMain.Size = new Size(mainDisplaySizeX, mainDisplaySizeY);
                    pictureBoxMain.Location = new System.Drawing.Point(0, 0);

                    //resets the previously saved points from moving
                    picMovedPositionX = 0;
                    picMovedPositionY = 0;
                    zoom = 1.0f;
                    CurrentZoom = 0;
                    pictureBoxMain.Invalidate();

                    //clear the highlight
                    pictureBoxMain.Controls.Clear();
                }
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }


        private void leftBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    if (!(pictureBoxMain.Location.X < 10))
                    {
                        pictureBoxMain.Location = new Point(pictureBoxMain.Location.X - 5, currentPicBoxLocation.Y);
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }
        private void rightBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    pictureBoxMain.Location = new Point(pictureBoxMain.Location.X + 5, pictureBoxMain.Location.Y);
                    currentPicBoxLocation = pictureBoxMain.Location;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void upBtn_MouseDown(object sender, MouseEventArgs e)
        {
            timer2.Enabled = true;
            timer2.Start();
        }


        private void upBtn_MouseUp(object sender, MouseEventArgs e)
        {
            timer2.Stop();
        }

        private void downBtn_MouseDown(object sender, MouseEventArgs e)
        {

            timer1.Enabled = true;
            timer1.Start();
        }


        private void downBtn_MouseUp(object sender, MouseEventArgs e)
        {
            timer1.Stop();
        }

        private void downBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    pictureBoxMain.Location = new Point(pictureBoxMain.Location.X, pictureBoxMain.Location.Y + 5);
                    currentPicBoxLocation = pictureBoxMain.Location;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void upBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {

                    if (!(pictureBoxMain.Location.Y < 50))
                    {
                        pictureBoxMain.Location = new Point(pictureBoxMain.Location.X, currentPicBoxLocation.Y - 5);
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            pictureBoxMain.Location = new Point(pictureBoxMain.Location.X, pictureBoxMain.Location.Y + 5);

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            pictureBoxMain.Location = new Point(pictureBoxMain.Location.X, pictureBoxMain.Location.Y - 5);

        }

        private void fitBtn_Click(object sender, EventArgs e)
        {
            try
            {
                pictureBoxMain.Controls.Clear();  
                int pbWidth = pictureBoxMain.Size.Width;
                int panelWidth = splitContainer2.Size.Width;
                int pbHeight = pictureBoxMain.Size.Height;
                int panelHeight = splitContainer2.Size.Height - 200;
                int count = 0;
                int count2 = 0;
                int counter = 0;


                if (pbWidth > panelWidth)
                {
                    count += (pbWidth - panelWidth) / 100;
                }
                if (pbHeight > panelHeight)
                {
                    count2 += (pbHeight - panelHeight) / 100;
                }
                if (count > count2)
                {
                    counter = count;
                }
                else
                {
                    counter = count2;
                }

                for (int i = 0; i <= counter; i++)
                {
                    pbWidth = pictureBoxMain.Size.Width;
                    pbHeight = pictureBoxMain.Size.Height;

                    if (pbWidth > panelWidth || pbHeight > panelHeight)
                    {
                        zoom -= 0.1f;
                        pictureBoxMain.Size = new Size((int)(pbOriginalWidth * zoom), (int)(pbOriginalHeight * zoom));
                    }
                    else
                    {
                        break;
                    }

                }

                pbWidth = pictureBoxMain.Size.Width;
                pbHeight = pictureBoxMain.Size.Height;
                pictureBoxMain.Location = new System.Drawing.Point((panelWidth - pbWidth) / 2, (panelHeight - pbHeight) / 2);
                pictureBoxMain.Invalidate();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }
        private void changeDetailForm(int index)
        {
            if (!animationFrmXl[index].Equals(""))
            {
                string folder = fpath;
                string str = folder.Replace(filename, "");
                string str2 = str + "\\animation\\" + animationFrmXl[index];
                richTextPath = str2;
            }
            else
            {
                richTextPath = "";

            }

            if (!form.IsDisposed)
            {

                //set detailform animation          
                form.path_RTF = richTextPath;
                //form.Size = new Size
                detailFormLocation = form.Location;

            }
        }
        private void setDetailPanel()
        {
            if (richTextPath.Equals(""))
            {
                animationTextBox.Clear();

            }
            else
            {
                animationTextBox.LoadFile(richTextPath);

            }

        }

        private void ExcelDataGrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                ExcelDataGrid.BeginEdit(true);
                fileSaved = false;
                //checkbox control
                if (e.ColumnIndex == 0)
                {
                    if (typeFrmXl[ExcelDataGrid.CurrentRow.Index].Equals("group"))
                    {

                        for (int i = selectedMainGroupIndex; i <= selectedMainEndGroupIndex; i++)
                        {
                            ExcelDataGrid[0, i].Value = true;
                        }
                    }
                }
                ExcelDataGrid.EndEdit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        public void UpdateExcelUsingEPPlus()
        {
            if (excelLoaded)
            {
                ExcelDataGrid.EndEdit();
                ExcelWorksheet myWorksheet = package.Workbook.Worksheets[dt.ToString()];
                myWorksheet.Cells["A1"].LoadFromDataTable(dt, true);
                package.Save();
                fileSaved = true;
            }
        }

        private void saveExcelBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    UpdateExcelUsingEPPlus();
                    statusLbl.Text = "Excel Sheet is saved.";
                    saveTimer.Enabled = true;
                    saveTimer.Start();
                    saveTimer.Interval = 3000;

                }

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }
        private void saveTimer_Tick(object sender, EventArgs e)
        {
            statusLbl.Text = "Ready";
            saveTimer.Stop();
        }
        private void timer3_Tick(object sender, EventArgs e)
        {
            timerIndex++;

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateExcelUsingEPPlus();
                statusLbl.Text = "Excel Sheet is saved.";
                saveTimer.Enabled = true;
                saveTimer.Start();
                saveTimer.Interval = 3000;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }


        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    if (!fileSaved)
                    {
                        //save the file first
                        DialogResult dialogResult = MessageBox.Show("Do you want to save the current file first?", "Save", MessageBoxButtons.YesNo);

                        if (dialogResult == DialogResult.Yes)
                        {
                            //toolStripProgressBar1.Visible = true;
                            UpdateExcelUsingEPPlus();
                            statusLbl.Text = "File is saved";
                            saveTimer.Enabled = true;
                            saveTimer.Start();
                            saveTimer.Interval = 3000;
                            //toolStripProgressBar1.Visible = false;

                        }
                    }

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "Save As";
                    saveFileDialog1.InitialDirectory = "";
                    saveFileDialog1.Filter = "Excel Files|*.xlsx";
                    saveFileDialog1.FilterIndex = 2;
                    saveFileDialog1.RestoreDirectory = true;

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string path = saveFileDialog1.FileName;
                        Stream strm = File.Create(path);
                        package.SaveAs(strm);
                        strm.Close();

                        statusLbl.Text = "New File is created and saved at " + path;
                    }

                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            Thread.Sleep(1000);

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBoxMain.Enabled = true;

        }

        void printDocument1_EndPrint(object sender, PrintEventArgs e)
        {
            currentPage = 0;
        }

        void printDocumentAll_EndPrint(object sender, PrintEventArgs e)
        {
            currentPage = 0;
        }

        void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            int checkedCount = pForm.checkedListBox1.CheckedItems.Count;
            Image img = Image.FromFile(filePaths[checkedIndices[currentPage]]);

            String str = Path.GetFileNameWithoutExtension(filePaths[checkedIndices[currentPage]]);
            Font drawFont = new Font("Arial", 16);
            SolidBrush drawBrush = new SolidBrush(Color.Black);
            PointF drawPoint = new Point(20, 20);

            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.DrawString(str, drawFont, drawBrush, drawPoint);

            if (img.Width < chosenPaperSizeWidth && img.Height < chosenPaperSizeHeight)
            {
                g.DrawImage(img, new Rectangle(marginX, marginY, img.Width, img.Height));

            }
            else
            {
                //try to fit the image to the width of the paper        
                //if landscape width becomes height so fit to height
                if (landscape)
                {
                    int fitLandscapeWidthX = img.Width * (chosenPaperSizeHeight - 2 * marginX) / img.Width;
                    int fitLandscapeWidthY = img.Height * (chosenPaperSizeHeight - 2 * marginX) / img.Width;

                    int fitLandscapeHeightX = img.Width * (chosenPaperSizeWidth - (marginY + marginX)) / img.Height;
                    int fitLandscapeHeightY = img.Height * (chosenPaperSizeWidth - (marginY + marginX)) / img.Height;

                    //if newly scaled image height is greater than the paper height, it will be cut off
                    if (fitLandscapeWidthY + marginY + marginY - 40 > chosenPaperSizeWidth)
                    {
                        g.DrawImage(img, new Rectangle(marginX, marginY, fitLandscapeHeightX, fitLandscapeHeightY));

                    }
                    else
                    {
                        g.DrawImage(img, new Rectangle(marginX, marginY, fitLandscapeWidthX, fitLandscapeWidthY));

                    }
                }
                else
                {
                    int fitPaperWidthX = img.Width * (chosenPaperSizeWidth - 2 * marginX) / img.Width;
                    int fitPaperWidthY = img.Height * (chosenPaperSizeWidth - 2 * marginX) / img.Width;
                    g.DrawImage(img, new Rectangle(marginX, marginY, fitPaperWidthX, fitPaperWidthY));
                }
            }

            currentPage++;
            e.HasMorePages = currentPage < checkedCount;
        }

        void printDocumentAll_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                Image img = Image.FromFile(filePaths[currentPage]);

                string str = Path.GetFileNameWithoutExtension(filePaths[currentPage]);
                Font drawFont = new Font("Arial", 16);
                SolidBrush drawBrush = new SolidBrush(Color.Black);
                PointF drawPoint = new Point(20, 20);

                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawString(str, drawFont, drawBrush, drawPoint);

                //try to fit the image to the width of the paper        
                //if landscape width becomes height so fit to height
                if (landscape)
                {
                    int fitLandscapeWidthX = mainDisplaySizeX * (chosenPaperSizeHeight - 2 * marginX) / mainDisplaySizeX;
                    int fitLandscapeWidthY = mainDisplaySizeY * (chosenPaperSizeHeight - 2 * marginX) / mainDisplaySizeX;

                    int fitLandscapeHeightX = mainDisplaySizeX * (chosenPaperSizeWidth - (marginY + marginX)) / mainDisplaySizeY;
                    int fitLandscapeHeightY = mainDisplaySizeY * (chosenPaperSizeWidth - (marginY + marginX)) / mainDisplaySizeY;

                    //if newly scaled image height is greater than the paper height, it will be cut off
                    if (fitLandscapeWidthY + marginY + marginY - 40 > chosenPaperSizeWidth)
                    {
                        g.DrawImage(img, new Rectangle(marginX, marginY, fitLandscapeHeightX, fitLandscapeHeightY));

                    }
                    else
                    {
                        g.DrawImage(img, new Rectangle(marginX, marginY, fitLandscapeWidthX, fitLandscapeWidthY));

                    }
                }
                else
                {
                    int fitPaperWidthX = mainDisplaySizeX * (chosenPaperSizeWidth - 2 * marginX) / mainDisplaySizeX;
                    int fitPaperWidthY = mainDisplaySizeY * (chosenPaperSizeWidth - 2 * marginX) / mainDisplaySizeX;
                    g.DrawImage(img, new Rectangle(marginX, marginY, fitPaperWidthX, fitPaperWidthY));
                }

                currentPage++;
                e.HasMorePages = currentPage < filePaths.Length;

                img = null;
                drawFont = null;
                drawBrush = null;
            }
            catch (Exception er)
            {
                MessageBox.Show(er.ToString());
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void pForm_printForm_Load(object sender, EventArgs e)
        {

            if (pForm.checkedListBox1.CheckedItems.Count == 0)
            {
                pForm.printBtn.Enabled = false;
            }

        }

        private void pForm_printBtn_Click(object sender, EventArgs e)
        {
            try
            {
                int count = 0;
                int checkedCount = pForm.checkedListBox1.CheckedItems.Count;
                //exclude select all check box
                int checkBoxCount = pForm.checkedListBox1.Items.Count;
                checkedFiles = new string[checkedCount];
                checkedIndices = new int[checkedCount];

                for (int j = 0; j < checkBoxCount; j++)
                {
                    if (pForm.checkedListBox1.GetItemChecked(j) == true)
                    {
                        checkedFiles[count] = pForm.checkedListBox1.Items[j].ToString();
                        count++;
                    }
                }

                //store new values to checkedindices array
                for (int i = 0; i < filePaths.Length; i++)
                {
                    for (int k = 0; k < checkedCount; k++)
                    {
                        if (filePaths[i].Contains(checkedFiles[k]))
                        {
                            checkedIndices[k] = i;
                        }
                    }
                }
                printSelectedPages();
                pForm.Close();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void printPreviewBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    pForm = new printForm();
                    pForm.printBtn.Click += this.pForm_printBtn_Click;
                    pForm.checkedListBox1.SelectedValueChanged += this.pForm_checkedListBox1_SelectedValueChanged;
                    pForm.checkBoxAll.CheckedChanged += this.pForm_checkBoxAll_CheckedChanged;
                    pForm.Load += this.pForm_printForm_Load;

                    var items = pForm.checkedListBox1.Items;

                    for (int i = 0; i < sheetNames.Length; i++)
                    {
                        if (sheetNames[i].Contains("Grp"))
                        {
                            items.Add(sheetNames[i]);
                        }
                    }
                    pForm.Show();
                }

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }
        private void pForm_checkedListBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (pForm.checkedListBox1.CheckedItems.Count == 0)
            {
                pForm.printBtn.Enabled = false;
            }
            else
            {
                pForm.printBtn.Enabled = true;
            }
        }
        private void pForm_checkBoxAll_CheckedChanged(object sender, EventArgs e)
        {
            if (pForm.checkBoxAll.Checked)
            {
                SelectDeselectAll(true);
                pForm.printBtn.Enabled = true;
            }
            else
            {
                SelectDeselectAll(false);
                pForm.printBtn.Enabled = false;
            }
        }
        private void SelectDeselectAll(bool bSelected)
        {
            for (int i = 0; i < pForm.checkedListBox1.Items.Count; i++)
            {
                if (bSelected)
                {
                    pForm.checkedListBox1.SetItemChecked(i, true);

                }
                else
                {
                    pForm.checkedListBox1.SetItemChecked(i, false);
                }
            }
        }
        public void printSelectedPages()
        {
            currentPage = 0;
            PrintDocument printDocument1 = new PrintDocument();
            PrintDialog myPrintDialog = new PrintDialog();

            preview = new PrintPreviewDialog();
            preview.Icon = Resources.Graphicloads_Polygon_Tools;

            printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printDocument1_PrintPage);
            printDocument1.EndPrint += new System.Drawing.Printing.PrintEventHandler(printDocument1_EndPrint);

            myPrintDialog.Document = printDocument1;

            if (myPrintDialog.ShowDialog() == DialogResult.OK)
            {
                chosenPaperSizeHeight = myPrintDialog.PrinterSettings.DefaultPageSettings.PaperSize.Height;
                chosenPaperSizeWidth = myPrintDialog.PrinterSettings.DefaultPageSettings.PaperSize.Width;
                landscape = myPrintDialog.PrinterSettings.DefaultPageSettings.Landscape;

                preview.WindowState = FormWindowState.Maximized;
                preview.Document = printDocument1;
                preview.ShowDialog();
            }

            printDocument1.Dispose();
        }

        public void printAllPages()
        {
            currentPage = 0;
            PrintDocument printDocumentAll = new PrintDocument();
            PrintDialog printDialog = new PrintDialog();

            printDialog.Document = printDocumentAll;

            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                chosenPaperSizeHeight = printDialog.PrinterSettings.DefaultPageSettings.PaperSize.Height;
                chosenPaperSizeWidth = printDialog.PrinterSettings.DefaultPageSettings.PaperSize.Width;
                landscape = printDialog.PrinterSettings.DefaultPageSettings.Landscape;
                printDocumentAll.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printDocumentAll_PrintPage);
                printDocumentAll.Print();
                statusLbl.Text = "Printing Done.";
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void printBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded) printAllPages();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void searchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                searchBtn.PerformClick();
            }
        }

        //search next file that contains the search value
        private void searchBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    //clear highlights make this into a method
                    animationTextBox.Clear();
                    clearGridHighlight();
                    ExcelDataGrid.ClearSelection();
                    //loop thru only rows that have animation
                    for (int i = 0; i < rowCount; i++)
                    {
                        if (searchBox.Text.Equals(""))
                        {
                            statusLbl.Text = "Alert: Enter Search Value.";
                            break;
                        }
                        else
                        {
                            //if name matches search value
                            if (nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase))
                            {
                                ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[i].Cells[0];
                                ExcelDataGrid.Rows[i].Selected = true;
                                gridHighlight();
                                changeDetailForm(i);
                                setDetailPanel();
                                statusLbl.Text = "Row " + i + " has matching results.";
                                countSearchFiles = i;
                                highlight_OnClick(i);
                                break;
                            }

                            // if name does not match search value but it has animation that matches the search result
                            else if (!nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !animationFrmXl[i].Equals(""))
                            {
                                string folder = fpath;
                                string str = folder.Replace(filename, "");
                                string str2 = str + "\\animation\\" + animationFrmXl[i];
                                string richTextPath2 = str2;

                                RichTextBox tempBox = new RichTextBox();
                                tempBox.LoadFile(richTextPath2);

                                if (tempBox.Text.CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase))
                                {
                                    ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[i].Cells[0];
                                    ExcelDataGrid.Rows[i].Selected = true;
                                    gridHighlight();
                                    changeDetailForm(i);
                                    setDetailPanel();
                                    statusLbl.Text = "Row " + i + " has matching results.";
                                    countSearchFiles = i;
                                    highlight_OnClick(i);
                                    break;
                                }
                            }
                            else
                            {
                                if (i == rowCount - 1)
                                {
                                    MessageBox.Show("Alert: Searched throught all the files.");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void searchPrvBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    //clear highlights
                    if (searchAllClicked)
                    {
                        pictureBoxMain.Invalidate();
                        searchAllClicked = false;
                    }
                    animationTextBox.Clear();
                    clearGridHighlight();
                    ExcelDataGrid.ClearSelection();

                    if (!searchingTextbox)
                    {
                        if (countSearchFiles < 2)
                        {
                            countSearchFiles = 2;
                        }
                        //loop thru only rows that have animation
                        for (int i = countSearchFiles - 2; i >= 0; i--)
                        {
                            //if name matches search value
                            if (nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !searchBox.Text.Equals(""))
                            {
                                ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[i].Cells[0];
                                ExcelDataGrid.Rows[i].Selected = true;
                                gridHighlight();
                                changeDetailForm(i);
                                setDetailPanel();
                                countSearchFiles = i + 1;
                                statusLbl.Text = "Row " + i + " has matching results.";
                                highlight_OnClick(i);
                                //if name matches search value and it has animation
                                if (!animationFrmXl[i].Equals(""))
                                {
                                    searchingTextbox = true;
                                }
                                break;
                            }

                            // if name does not match search value but it has animation that matches the search result
                            else if (!nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !animationFrmXl[i].Equals(""))
                            {
                                string folder = fpath;
                                string str = folder.Replace(filename, "");
                                string str2 = str + "\\animation\\" + animationFrmXl[i];
                                string richTextPath2 = str2;

                                RichTextBox tempBox = new RichTextBox();
                                tempBox.LoadFile(richTextPath2);

                                if (tempBox.Text.CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !searchBox.Text.Equals(""))
                                {
                                    ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[i].Cells[0];
                                    ExcelDataGrid.Rows[i].Selected = true;
                                    gridHighlight();
                                    changeDetailForm(i);
                                    setDetailPanel();
                                    countSearchFiles = i + 1;
                                    statusLbl.Text = "Row " + i + " has matching results.";
                                    highlight_OnClick(i);
                                    searchingTextbox = true;
                                    break;
                                }
                            }
                            else
                            {
                                if (i == 0)
                                {
                                    MessageBox.Show("Searched through all the files.");
                                    statusLbl.Text = "Ready";
                                    countSearchFiles = rowCount - 1;
                                    resetSearchFiles = true;
                                }
                            }
                        }
                    }

                    if (!resetSearchFiles)
                    {
                        //when user clicks btn again it loses focus, so need to call this line again
                        ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[countSearchFiles - 1].Cells[0];
                        ExcelDataGrid.Rows[countSearchFiles - 1].Selected = true;
                        gridHighlight();
                        changeDetailForm(countSearchFiles - 1);
                        setDetailPanel();

                        if (searchedAllText)
                        {
                            animationTextBox.SelectAll();
                            animationTextBox.SelectionBackColor = Color.White;
                            animationTextBox.DeselectAll();
                            searchedAllText = false;
                        }

                        if (searchBox.Text.Equals(""))
                        {
                            statusLbl.Text = "No results Found";
                        }
                        else
                        {
                            // for detail form
                            if (form.Visible && !animationTextBox.Text.Equals(""))
                            {
                                form.animationTextBox.SelectionBackColor = Color.White;
                                selectionStart2 = form.animationTextBox.Find(searchBox.Text, 0, selectionStart2, RichTextBoxFinds.Reverse);
                                form.animationTextBox.ScrollToCaret();
                                selectionStop2 = selectionStart2 + searchBox.Text.Length;

                                if (selectionStart2 == -1)
                                {
                                    searchingTextbox = false;
                                    selectionStart = 0;
                                }
                                form.animationTextBox.SelectionBackColor = Color.Yellow;
                            }
                            else if (!animationTextBox.Text.Equals(""))
                            {
                                animationTextBox.SelectionBackColor = Color.White;
                                selectionStart = animationTextBox.Find(searchBox.Text, 0, selectionStart, RichTextBoxFinds.Reverse);
                                animationTextBox.ScrollToCaret();
                                selectionStop = selectionStart + searchBox.Text.Length;

                                if (selectionStart == -1)
                                {
                                    searchingTextbox = false;
                                    selectionStart = 0;
                                }
                                animationTextBox.SelectionBackColor = Color.Yellow;
                            }
                            else
                            {
                                searchingTextbox = false;
                            }

                        }
                    }
                    else
                    {
                        resetSearchFiles = false;

                    }
                }

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void findAllBtn_Click(object sender, EventArgs e)
        {
            try
            {
                searchedAllText = true;

                if (searchBox.Text.Equals(""))
                {
                    statusLbl.Text = "No results Found";
                }
                else
                {
                    animationTextBox.SelectionStart = 0;
                    animationTextBox.SelectionLength = animationTextBox.Text.Length;
                    animationTextBox.SelectionBackColor = Color.White;

                    int index = 0;
                    int countResults = 0;

                    while (index < animationTextBox.Text.LastIndexOf(searchBox.Text, StringComparison.InvariantCultureIgnoreCase))
                    {
                        animationTextBox.Find(searchBox.Text, index, animationTextBox.TextLength, RichTextBoxFinds.None);
                        animationTextBox.SelectionBackColor = Color.Yellow;
                        index = animationTextBox.Text.IndexOf(searchBox.Text, index, StringComparison.InvariantCultureIgnoreCase) + 1;
                        countResults = new Regex(searchBox.Text).Matches(animationTextBox.Text).Count;
                    }
                    statusLbl.Text = countResults + " results Found";
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void searchAllBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    searchAllClicked = true;
                    animationTextBox.Clear();
                    clearGridHighlight();
                    removeBtn.PerformClick();
                    ExcelDataGrid.ClearSelection();
                    for (int i = 0; i < rowCount; i++)
                    {
                        if (nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !searchBox.Text.Equals(""))
                        {
                            ExcelDataGrid.Rows[i].Selected = true;
                            multipleHighlight_OnClick(i);
                        }
                        else if (!nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !animationFrmXl[i].Equals(""))
                        {
                            string folder = fpath;
                            string str = folder.Replace(filename, "");
                            string str2 = str + "\\animation\\" + animationFrmXl[i];
                            string richTextPath2 = str2;

                            RichTextBox tempBox = new RichTextBox();
                            tempBox.LoadFile(richTextPath2);


                            if (tempBox.Text.CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !searchBox.Text.Equals(""))
                            {
                                ExcelDataGrid.Rows[i].Selected = true;
                                multipleHighlight_OnClick(i);
                            }

                        }

                    }

                }

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void searchNextBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    //clear highlights
                    if (searchAllClicked)
                    {
                        pictureBoxMain.Invalidate();
                        searchAllClicked = false;
                    }
                    animationTextBox.Clear();
                    clearGridHighlight();
                    ExcelDataGrid.ClearSelection();

                    if (!searchingTextbox)
                    {
                        //loop thru only rows that have animation
                        for (int i = countSearchFiles; i < rowCount; i++)
                        {
                            //if name matches search value
                            if (nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !searchBox.Text.Equals(""))
                            {

                                ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[i].Cells[0];
                                ExcelDataGrid.Rows[i].Selected = true;
                                gridHighlight();
                                changeDetailForm(i);
                                setDetailPanel();
                                countSearchFiles = i + 1;
                                statusLbl.Text = "Row " + i + " has matching results.";
                                highlight_OnClick(i);
                                //if name matches search value and it has animation
                                if (!animationFrmXl[i].Equals(""))
                                {
                                    searchingTextbox = true;
                                }
                                break;
                            }

                            // if name does not match search value but it has animation that matches the search result
                            else if (!nameFrmXl[i].CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !animationFrmXl[i].Equals(""))
                            {
                                string folder = fpath;                               
                                string richTextPath2 = folder.Replace(filename, "") + "\\animation\\" + animationFrmXl[i]; ;

                                RichTextBox tempBox = new RichTextBox();
                                tempBox.LoadFile(richTextPath2);

                                if (tempBox.Text.CaseInsensitiveContains(searchBox.Text, StringComparison.OrdinalIgnoreCase) && !searchBox.Text.Equals(""))
                                {
                                    ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[i].Cells[0];
                                    ExcelDataGrid.Rows[i].Selected = true;
                                    gridHighlight();
                                    changeDetailForm(i);
                                    setDetailPanel();
                                    countSearchFiles = i + 1;
                                    statusLbl.Text = "Row " + i + " has matching results.";
                                    highlight_OnClick(i);
                                    searchingTextbox = true;
                                    break;
                                }
                            }
                            else
                            {
                                if (i == rowCount - 1)
                                {
                                    MessageBox.Show("Searched through all the files.");
                                    statusLbl.Text = "Ready";
                                    countSearchFiles = 0;
                                    resetSearchFiles = true;
                                }
                            }
                        }
                    }

                    if (!resetSearchFiles)
                    {
                        //when user clicks btn again it loses focus, so need to call this line again
                        ExcelDataGrid.CurrentCell = ExcelDataGrid.Rows[countSearchFiles - 1].Cells[0];
                        ExcelDataGrid.Rows[countSearchFiles - 1].Selected = true;
                        gridHighlight();
                        changeDetailForm(countSearchFiles - 1);
                        setDetailPanel();

                        if (searchedAllText)
                        {
                            animationTextBox.SelectAll();
                            animationTextBox.SelectionBackColor = Color.White;
                            animationTextBox.DeselectAll();
                            searchedAllText = false;
                        }

                        if (searchBox.Text.Equals(""))
                        {
                            statusLbl.Text = "No results Found";
                        }
                        else
                        {
                            // for detail form
                            if (form.Visible && !animationTextBox.Text.Equals(""))
                            {
                                form.animationTextBox.SelectionBackColor = Color.White;
                                selectionStart2 = form.animationTextBox.Find(searchBox.Text, selectionStop2, form.animationTextBox.TextLength, RichTextBoxFinds.None);
                                form.animationTextBox.ScrollToCaret();
                                selectionStop2 = selectionStart2 + searchBox.Text.Length;

                                if (selectionStart2 == -1)
                                {
                                    searchingTextbox = false;
                                    //reset for different files
                                    selectionStart = 0;
                                }
                                form.animationTextBox.SelectionBackColor = Color.Yellow;
                            }
                            else if (!animationTextBox.Text.Equals(""))
                            {

                                animationTextBox.SelectionBackColor = Color.White;
                                selectionStart = animationTextBox.Find(searchBox.Text, selectionStop, animationTextBox.TextLength, RichTextBoxFinds.None);
                                animationTextBox.ScrollToCaret();
                                selectionStop = selectionStart + searchBox.Text.Length;

                                if (selectionStart == -1)
                                {
                                    searchingTextbox = false;
                                    selectionStart = 0;
                                }
                                animationTextBox.SelectionBackColor = Color.Yellow;
                            }
                            else
                            {
                                searchingTextbox = false;
                            }
                        }
                    }
                    else
                    {
                        resetSearchFiles = false;
                    }
                }

            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }

        }

        private void saveScreenshots()
        {
            try
            {
                for (int c = 0; c < sheetNames.Length; c++)
                {
                    if (sheetNames[c].Contains("Grp"))
                    {
                        // pick the table using the name of the sheet
                        dt = result.Tables[c];

                        //remove first row to make it easier to convert
                        rowCount = dt.Rows.Count;
                        colCount = dt.Columns.Count;

                        typeFrmXl = new string[rowCount];
                        foreColorFrmXl = new string[rowCount];
                        backColorFrmXl = new string[rowCount];
                        xPositionFrmXl = new int[rowCount];
                        yPositionFrmXl = new int[rowCount];
                        xSizeFrmXl = new int[rowCount];
                        ySizeFrmXl = new int[rowCount];
                        nameFrmXl = new string[rowCount];
                        configFrmXl = new string[rowCount];
                        animationFrmXl = new string[rowCount];

                        for (int i = 0; i < rowCount; i++)
                        {
                            typeFrmXl[i] = dt.Rows[i][1].ToString();
                            nameFrmXl[i] = dt.Rows[i][2].ToString();
                            foreColorFrmXl[i] = dt.Rows[i][3].ToString();
                            backColorFrmXl[i] = dt.Rows[i][4].ToString();
                            xPositionFrmXl[i] = Convert.ToInt32(dt.Rows[i][5]);
                            yPositionFrmXl[i] = Convert.ToInt32(dt.Rows[i][6]);
                            xSizeFrmXl[i] = Convert.ToInt32(dt.Rows[i][7]);
                            ySizeFrmXl[i] = Convert.ToInt32(dt.Rows[i][8]);
                            configFrmXl[i] = dt.Rows[i][9].ToString();
                            animationFrmXl[i] = dt.Rows[i][10].ToString();
                        }

                        //pictureBoxMain.Visible = true;
                        pictureBoxMain.Enabled = true;
                        excelLoaded = true;

                        pictureBoxMain.Invalidate();
                        Bitmap bmp = new Bitmap(xSizeFrmXl[0], ySizeFrmXl[0]);
                        pictureBoxMain.DrawToBitmap(bmp, new Rectangle(0, 0, xSizeFrmXl[0], ySizeFrmXl[0]));

                        string path = fpath.Replace(filename, "");

                        if (!Directory.Exists(path + "\\screenshots\\"))
                        {
                            Directory.CreateDirectory(path + "\\screenshots\\");

                        }
                        bmp.Save(path + "\\screenshots\\" + sheetNames[c] + ".png", ImageFormat.Png);
                    }
                }
            }
            catch (Exception error)
            {
                error.ToString();
            }
        }

        private void reCenterPosition()
        {
            int xPosi = xPositionFrmXl[ExcelDataGrid.CurrentRow.Index];
            int yPosi = yPositionFrmXl[ExcelDataGrid.CurrentRow.Index];
            int newPointX = 0;
            int newPointY = 0;

            if (xPosi * zoom > (0.9 * splitContainer2.Size.Width) && yPosi * zoom > (0.7 * splitContainer2.Size.Height))
            {
                newPointX = currentPicBoxLocation.X - xPosi / 2;
                newPointY = currentPicBoxLocation.Y - yPosi / 2;
                pictureBoxMain.Location = new Point((int)(newPointX * zoom), (int)(newPointY * zoom));
            }
            else if (xPosi * zoom > (0.9 * splitContainer2.Size.Width))
            {
                newPointX = currentPicBoxLocation.X - xPosi / 2;
                newPointY = currentPicBoxLocation.Y - yPosi / 2;

                pictureBoxMain.Location = new Point((int)(newPointX * zoom), (int)(newPointY * zoom));

            }
            else if (yPosi * zoom > (0.7 * splitContainer2.Size.Height))
            {
                newPointX = currentPicBoxLocation.X - xPosi / 2;
                newPointY = currentPicBoxLocation.Y - yPosi / 2;
                pictureBoxMain.Location = new Point((int)(newPointX * zoom), (int)(newPointY * zoom));
            }

            else
            {
                pictureBoxMain.Location = currentPicBoxLocation;
            }
        }

        private void clearGridHighlight()
        {
            for (int c = 0; c < rowCount; c++)
            {
                ExcelDataGrid.Rows[c].DefaultCellStyle.BackColor = Color.White;


            }
        }

        private void gridHighlight()
        {
            if (!highlightChecked)
            {
                clearGridHighlight();

                if (ExcelDataGrid.CurrentRow.Selected)
                {

                    currentSelectedRow = ExcelDataGrid.CurrentRow.Index;
                    selectedGroupIndex = 0;
                    selectedMainGroupIndex = 0;
                    selectedEndGroupIndex = 0;
                    selectedMainEndGroupIndex = 0;

                    int counter = 0;
                    int counter2 = 0;
                    int groupCount = 0;
                    int count = 0;

                    int endGroup = 0;
                    bool greater = false;
                    bool noGroup = false;
                    bool noEndGroup = false;


                    int groupCounter = 0;
                    int endGroupCounter = 0;
                    for (int k = currentSelectedRow; k >= 0; k--)
                    {
                        if (typeFrmXl[k].Equals("group"))
                        {
                            selectedMainGroupIndex = k;
                            groupCounter++;
                            break;
                        }
                        if ((typeFrmXl[k].CaseInsensitiveContains("endgroup") && !typeFrmXl[k].Contains("*endgroup")) || typeFrmXl[k].Contains("endgroup *endgroup"))
                        {
                            noEndGroup = true;
                            endGroupCounter++;
                        }


                    }

                    if (groupCounter == 0 && endGroupCounter == 0)
                    {
                        selectedGroupIndex = 0;
                        selectedMainGroupIndex = 0;
                        selectedEndGroupIndex = 0;
                        selectedMainEndGroupIndex = 0;
                        noEndGroup = true;
                        noGroup = true;


                    }

                    for (int n = currentSelectedRow; n < rowCount; n++)
                    {

                        if ((typeFrmXl[n].CaseInsensitiveContains("endgroup") && !typeFrmXl[n].Contains("*endgroup")) || typeFrmXl[n].Contains("endgroup *endgroup"))
                        {
                            selectedMainEndGroupIndex = n;
                            break;
                        }
                        if (typeFrmXl[n].Equals("group"))
                        {
                            noGroup = true;

                        }

                    }

                    // go to previous index to find group
                    for (int i = currentSelectedRow; i >= 0; i--)
                    {
                        if (typeFrmXl[i].Equals("*group"))
                        {

                            count++;

                            if (count == counter)
                            {
                                if (greater)
                                {
                                    selectedGroupIndex = i;
                                    break;

                                }
                                else
                                {
                                    selectedGroupIndex = i;
                                }


                            }
                            else if (counter > count)
                            {

                            }

                            else
                            {
                                selectedGroupIndex = i;
                                break;
                            }


                        }
                        else if (typeFrmXl[i].CaseInsensitiveContains("*endgroup"))
                        {
                            if (currentSelectedRow == i)
                            {
                                greater = true;
                            }
                            //endgroup counter
                            counter++;
                        }
                        else if (typeFrmXl[i].Equals("group"))
                        {
                            selectedGroupIndex = i;
                            break;
                        }
                        else if ((typeFrmXl[i].CaseInsensitiveContains("endgroup") && !typeFrmXl[i].Contains("*endgroup")) || typeFrmXl[i].Contains("endgroup *endgroup"))

                        {
                            {
                                selectedGroupIndex = selectedMainGroupIndex;
                                break;
                            }
                        }
                        else
                        {
                            endGroup++;
                        }
                    }

                    // start the loop from the group even if you clicked on an object.
                    for (int j = selectedGroupIndex; j < rowCount; j++)
                    {
                        if (typeFrmXl[j].Equals("*group"))
                        {
                            groupCount++;

                        }

                        else if (typeFrmXl[j].CaseInsensitiveContains("*endgroup"))
                        {
                            counter2++;

                            // if you click on a group
                            if (counter2 == groupCount)
                            {
                                selectedEndGroupIndex = j;
                                break;
                            }


                        }
                        else if (typeFrmXl[j].Equals("group"))
                        {
                            selectedEndGroupIndex = selectedMainEndGroupIndex;
                            break;
                        }
                        else if ((typeFrmXl[j].CaseInsensitiveContains("endgroup") && !typeFrmXl[j].Contains("*endgroup")) || typeFrmXl[j].Contains("endgroup *endgroup"))

                        {
                            {
                                selectedEndGroupIndex = selectedMainEndGroupIndex;
                                break;
                            }
                        }
                    }

                    if (noGroup && noEndGroup)
                    {

                    }
                    else
                    {
                        highlight_grid(selectedMainGroupIndex, selectedGroupIndex, selectedMainEndGroupIndex, selectedEndGroupIndex);

                    }
                }

            }
            else
            {
                clearGridHighlight();
            }
        }
        private void removeBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelLoaded)
                {
                    pictureBoxMain.Invalidate();

                    //clear the highlight
                    pictureBoxMain.Controls.Clear();
                }
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception error)
            {
                System.Windows.Forms.MessageBox.Show(error.ToString());
            }
        }

        private void ExcelDataGrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString());
        }

        private void fontUpBtn_Click(object sender, EventArgs e)
        {

            if (form.Visible)
            {
                form.animationTextBox.Font = new Font("Arial", form.animationTextBox.Font.Size + 1, FontStyle.Regular);
                this.CheckKeyword("property", Color.Purple, 0);
                this.CheckKeyword("name", Color.Green, 0);

            }
            else
            {
                animationTextBox.Font = new Font("Arial", animationTextBox.Font.Size + 1, FontStyle.Regular);
                this.CheckKeyword("property", Color.Purple, 0);
                this.CheckKeyword("name", Color.Green, 0);

            }

        }

        private void fontDownBtn_Click(object sender, EventArgs e)
        {
            if (form.Visible)
            {
                form.animationTextBox.Font = new Font("Arial", form.animationTextBox.Font.Size - 1, FontStyle.Regular);
                this.CheckKeyword("property", Color.Purple, 0);
                this.CheckKeyword("name", Color.Green, 0);

            }
            else
            {
                animationTextBox.Font = new Font("Arial", animationTextBox.Font.Size - 1, FontStyle.Regular);
                this.CheckKeyword("property", Color.Purple, 0);
                this.CheckKeyword("name", Color.Green, 0);
            }

        }

        private void animationTextBox_TextChanged(object sender, EventArgs e)
        {
            this.CheckKeyword("property", Color.Purple, 0);
            this.CheckKeyword("name", Color.Green, 0);
        }
        private void CheckKeyword(string word, Color color, int startIndex)
        {
            if (form.Visible)
            {
                if (form.animationTextBox.Text.CaseInsensitiveContains(word))
                {
                    int index = -1;
                    int selectStart = form.animationTextBox.SelectionStart;

                    while ((index = form.animationTextBox.Text.IndexOf(word, (index + 1))) != -1)
                    {
                        form.animationTextBox.Select((index + startIndex), word.Length);
                        form.animationTextBox.SelectionColor = color;
                        form.animationTextBox.Select(selectStart, 0);
                        form.animationTextBox.SelectionColor = Color.Black;
                    }
                }
            }
            else
            {
                if (this.animationTextBox.Text.CaseInsensitiveContains(word))
                {
                    int index = -1;
                    int selectStart = this.animationTextBox.SelectionStart;

                    while ((index = this.animationTextBox.Text.IndexOf(word, (index + 1))) != -1)
                    {
                        this.animationTextBox.Select((index + startIndex), word.Length);
                        this.animationTextBox.SelectionColor = color;
                        this.animationTextBox.Select(selectStart, 0);
                        this.animationTextBox.SelectionColor = Color.Black;
                    }
                }
            }
           
        }
        public void RotateRectangle(Graphics g, RectangleF r, float angle, Brush b)
        {
            try
            {
                using (Matrix m = new Matrix())
                {
                    m.RotateAt(90f, new PointF(r.Left + (r.Width / 2),
                                              r.Top + (r.Height / 2)));
                    g.Transform = m;
                    g.FillRectangle(b, r);
                    g.ResetTransform();
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            

        }
    }

    public static class Extensions
    {
        public static bool CaseInsensitiveContains(this string text, string value,
            StringComparison stringComparison = StringComparison.CurrentCultureIgnoreCase)
        {
            return text.IndexOf(value, stringComparison) >= 0;
        }
    }
}