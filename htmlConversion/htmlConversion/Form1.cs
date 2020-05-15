using ExCSS;
using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace htmlConversion
{
    public partial class Form1 : Form
    {

        RichTextBox richTB;
        bool mainGrpFlag = false;
        bool grpFlag = false;
        string mainGrpNodeParentId = "";
        string mainGrpNodeId = "";
        string grpNodeParentId = "";
        string grpNodeId = "";
        private string openFileName = "";
        string fileNamePure = "";
        private string openFilePath = "";
        private string fileNameExt = "";
        private int countObjs = 0;
        private int childCount = 0;
        private int nodeCount = 0;
        private int parentIndex = 0;
        HtmlNode previousNode = null;
        private string errorCheckNode = "";
        private string errorCheckFile = "";

        private string[] typeFrmXl;
        private string[] nameFrmXl;
        private string[] backColorFrmXl;
        private string[] foreColorFrmXl;
        private int[] xPositionFrmXl;
        private int[] yPositionFrmXl;
        private int[] xSizeFrmXl;
        private int[] ySizeFrmXl;
        private string[] configFrmXl;
        private string[] animationFrmXl;
        private string backgroundColorMD;
        private string widthMD;
        private string heightMD;
        string savePath;
        string saveFilePath;

        string rotation;
        string posiX_parent;
        string posiY_parent;
        string sizeX_parent;
        string sizeY_parent;

        private ExcelPackage package;
        private ExcelWorksheet ws;

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                package = new ExcelPackage();

                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Filter = "Html Files|*.htm;*.html";
                fdlg.Title = "Select a Html File";
                fdlg.Multiselect = true;
                //fileName = fdlg.SafeFileName;

                if (fdlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (String file in fdlg.FileNames)
                    {
                        //fileName = Path.GetFileNameWithoutExtension(fdlg.FileName);
                        openFilePath = Path.GetDirectoryName(file) + "/";
                        openFileName = Path.GetFileName(file);
                        listBox1.Items.Add(openFileName);
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
            }
        }

        private void createExcelFile(string file_name)
        {
            // errorCheckFile = file_name;
            ws = package.Workbook.Worksheets.Add("Grp_" + file_name);

            var headerRow = new List<string[]>()
                {
                  new string[] { "Check", "Object", "Name", "ForeColor", "BackColor", "PositionX",
                      "PositionY", "SizeX", "SizeY", "Configuration", "Animation" }
                };
            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

            ws.Cells[headerRange].LoadFromArrays(headerRow);
            ws.Cells[headerRange].Style.Font.Bold = true;
            ws.Cells[headerRange].Style.Font.Size = 14;
            ws.Cells[headerRange].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

            ws.Cells[2, 1].Value = "false";
            ws.Cells[2, 2].Value = "display";
            ws.Cells[2, 3].Value = "mainDisplay";
            ws.Cells[2, 4].Value = "";
            ws.Cells[2, 5].Value = backgroundColorMD;
            ws.Cells[2, 6].Value = 0;
            ws.Cells[2, 7].Value = 0;
            ws.Cells[2, 8].Value = Convert.ToInt32(widthMD.Replace("px", ""));
            ws.Cells[2, 9].Value = Convert.ToInt32(heightMD.Replace("px", ""));
            ws.Cells[2, 10].Value = "";
            ws.Cells[2, 11].Value = "";

            for (int i = 0; i < countObjs; i++)
            {
                ws.Cells[i + 3, 1].Value = "false";
                ws.Cells[i + 3, 2].Value = typeFrmXl[i];
                ws.Cells[i + 3, 3].Value = nameFrmXl[i];
                ws.Cells[i + 3, 4].Value = foreColorFrmXl[i];
                ws.Cells[i + 3, 5].Value = backColorFrmXl[i];
                ws.Cells[i + 3, 6].Value = xPositionFrmXl[i];
                ws.Cells[i + 3, 7].Value = yPositionFrmXl[i];
                ws.Cells[i + 3, 8].Value = xSizeFrmXl[i];
                ws.Cells[i + 3, 9].Value = ySizeFrmXl[i];
                ws.Cells[i + 3, 10].Value = configFrmXl[i];
                ws.Cells[i + 3, 11].Value = animationFrmXl[i];

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            richTB = new RichTextBox();

        }

        private void button1_Click(object sender, EventArgs e)
        {
           

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save As";
            saveFileDialog1.InitialDirectory = "";
            saveFileDialog1.Filter = "Excel Files|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                savePath =Path.GetDirectoryName(saveFileDialog1.FileName) + "\\";

                saveFilePath = saveFileDialog1.FileName;

                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    countObjs = 0;
                    convertFile(openFilePath + listBox1.Items[i].ToString());
                    createExcelFile(listBox1.Items[i].ToString());

                }

                Stream strm = File.Create(saveFilePath);
                package.SaveAs(strm);
                strm.Close();

                statusLbl.Text = "Saved";
            }
            saveTimer.Enabled = true;
            saveTimer.Start();
            saveTimer.Interval = 3000;
        }

        private void convertFile(string file)
        {
            fileNamePure = Path.GetFileNameWithoutExtension(file);

            var document = new HtmlAgilityPack.HtmlDocument();
            document.Load(file);

            //get info for main display
            var nodeMD = document.DocumentNode.SelectSingleNode("//div[@id='Background']");
            var parserMD = new Parser();
            var stylesheetMD = new StyleSheet();
            if (nodeMD.Attributes["style"] != null)
            {
                stylesheetMD = parserMD.Parse(".someClass{" + nodeMD.Attributes["style"].Value);
            }

            backgroundColorMD = stylesheetMD.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                 .SelectMany(r => r.Declarations)
                                                                 .FirstOrDefault(d => d.Name.CaseInsensitiveContains("color"))
                                                                 .Term
                                                                 .ToString();
            widthMD = stylesheetMD.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                             .SelectMany(r => r.Declarations)
                                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("width"))
                                             .Term
                                             .ToString();
            heightMD = stylesheetMD.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                             .SelectMany(r => r.Declarations)
                                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("height"))
                                             .Term
                                             .ToString();

            var node = document.DocumentNode.SelectSingleNode("//div[@id='Page']");

            foreach (var nodes in node.Descendants())
            {
                if (nodes.NodeType == HtmlNodeType.Element && nodes.HasAttributes)
                {
                    countObjs++;

                    //if (nodes.ChildNodes.Count > 1)
                    //{
                    //    //add twice of the these rows as group and end group 
                    //}
                }
                
            }
            int size = countObjs;
            typeFrmXl = new string[size];
            nameFrmXl = new string[size];
            foreColorFrmXl = new string[size];
            backColorFrmXl = new string[size];
            xPositionFrmXl = new int[size];
            yPositionFrmXl = new int[size];
            xSizeFrmXl = new int[size];
            ySizeFrmXl = new int[size];
            configFrmXl = new string[size];
            animationFrmXl = new string[size];

            int counter = 0;
            grpFlag = false;
            mainGrpFlag = false;
            try
            {
                foreach (var nodes in node.Descendants())
                {
                    if (nodes.NodeType == HtmlNodeType.Element && nodes.HasAttributes && nodes.Attributes["id"] != null)
                    {
                        var parser = new Parser();
                        var parser2 = new Parser();
                        var stylesheet = new StyleSheet();
                        var hdxProperty = new StyleSheet();

                        string extractFontSize = "";
                        string extractFont = "";
                        string extractFontStyle = "";
                        string extractForeColor = "";
                        string extractBackColor = "";

                        //MessageBox.Show(counter.ToString() + nodes.Id + nodes.Name + nodes.Line + nodes.InnerText + nodes.HasAttributes);
                        if (nodes.Attributes["style"] != null)
                        {
                            stylesheet = parser.Parse(".someClass{" + nodes.Attributes["style"].Value);
                        }
                        else if (nodes.Attributes["title style"] != null)
                        {
                            stylesheet = parser.Parse(".someClass{" + nodes.Attributes["title style"].Value);
                        }

                        if (nodes.Attributes["hdxproperties"] != null && nodes.Attributes["hdxproperties"].Value.CaseInsensitiveContains("textcolor"))
                        {
                            hdxProperty = parser2.Parse(".someClass{" + nodes.Attributes["hdxproperties"].Value);
                            //hdxProperty

                            extractBackColor = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("textcolor"))
                              .Term
                              .ToString();
                        }
                        else if (nodes.Attributes["hdxproperties"] != null && nodes.Attributes["hdxproperties"].Value.CaseInsensitiveContains("fillcolor"))
                        {
                            hdxProperty = parser2.Parse(".someClass{" + nodes.Attributes["hdxproperties"].Value);
                            //hdxProperty
                            extractBackColor = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor"))
                              .Term
                              .ToString();
                        }
                        else if (nodes.Attributes["hdxproperties"] != null && nodes.Attributes["hdxproperties"].Value.CaseInsensitiveContains("linecolor"))
                        {
                            hdxProperty = parser2.Parse(".someClass{" + nodes.Attributes["hdxproperties"].Value);
                            //hdxProperty
                            extractForeColor = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("linecolor"))
                              .Term
                              .ToString();
                        }

                        string extractPosiX = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                        .SelectMany(r => r.Declarations)
                        .FirstOrDefault(d => d.Name.Equals("LEFT", StringComparison.InvariantCultureIgnoreCase))
                        .Term
                        .ToString();

                        string extractPosiY = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                        .SelectMany(r => r.Declarations)
                        .FirstOrDefault(d => d.Name.Equals("TOP", StringComparison.InvariantCultureIgnoreCase))
                        .Term
                        .ToString();

                        string extractWidth = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                        .SelectMany(r => r.Declarations)
                        .FirstOrDefault(d => d.Name.Equals("Width", StringComparison.InvariantCultureIgnoreCase))
                        .Term
                        .ToString();

                        string extractHeight = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                        .SelectMany(r => r.Declarations)
                        .FirstOrDefault(d => d.Name.Equals("Height", StringComparison.InvariantCultureIgnoreCase))
                        .Term
                        .ToString();

                        if (nodes.Name.Contains("text") || nodes.Id.Contains("text") || nodes.Id.CaseInsensitiveContains("txt") || nodes.Id.Contains("question"))
                        {
                            //if font size not available, find its parent and use parent's font

                            if (!nodes.Attributes["style"].Value.CaseInsensitiveContains("font-family") && !nodes.Attributes["style"].Value.CaseInsensitiveContains("font-size"))
                            {
                                var stylesheetParent = parser.Parse(".someClass{" + nodes.ParentNode.Attributes["style"].Value);

                                extractFontSize = stylesheetParent.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-size"))
                                                                                   .Term
                                                                                   .ToString();
                                extractFontStyle = stylesheetParent.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-weigth"))
                                                                                   .Term
                                                                                   .ToString();
                                extractFont = stylesheetParent.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-family"))
                               .Term
                               .ToString();

                                string replaceSpace;

                                if (nodes.InnerText.CaseInsensitiveContains("nbsp"))
                                {
                                    string replaceNbsp = nodes.InnerText.Replace("&nbsp;", " ");

                                    replaceSpace = Regex.Replace(replaceNbsp, @"\s+", " ");
                                    configFrmXl[counter] = extractFont + "(||)" + extractFontSize.Replace("pt", "") + "(||)" + extractFontStyle + "(||)" + replaceSpace;
                                }
                                else
                                {
                                    replaceSpace = Regex.Replace(nodes.InnerText, @"\s+", " ");
                                    configFrmXl[counter] = extractFont + "(||)" + extractFontSize.Replace("pt", "") + "(||)" + extractFontStyle + "(||)" + replaceSpace;
                                }
                            }
                            else
                            {
                                //errorCheckNode = nodes.Id + nodes.Attributes["style"].Value;

                                extractFontSize = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-size"))
                                                                                   .Term
                                                                                   .ToString();
                                if(stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-weight")) != null)
                                {
                                    extractFontStyle = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-weight"))
                                                                                   .Term
                                                                                   .ToString();
                                }
                                else
                                {
                                    extractFontStyle = "normal";
                                }
                                

                                if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-family")) == null)
                                {
                                    var stylesheetParent = parser.Parse(".someClass{" + nodes.ParentNode.Attributes["style"].Value);

                                    extractFont = stylesheetParent.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-family"))
                               .Term
                               .ToString();
                                }
                                else
                                {
                                    if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                      .SelectMany(r => r.Declarations)
                                      .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-family")).Term == null)
                                    {
                                        extractFont = "Arial";
                                    }
                                    else
                                    {
                                        extractFont = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                       .SelectMany(r => r.Declarations)
                                       .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-family"))
                                       .Term
                                       .ToString();
                                    }

                                }


                                string replaceSpace;

                                if (nodes.InnerText.CaseInsensitiveContains("nbsp"))
                                {
                                    string replaceNbsp = nodes.InnerText.Replace("&nbsp;", " ");

                                    replaceSpace = Regex.Replace(replaceNbsp, @"\s+", " ");
                                    configFrmXl[counter] = extractFont + "(||)" + extractFontSize.Replace("pt", "") + "(||)" + extractFontStyle + "(||)" + replaceSpace;
                                }
                                else
                                {
                                    replaceSpace = Regex.Replace(nodes.InnerText, @"\s+", " ");
                                    configFrmXl[counter] = extractFont + "(||)" + extractFontSize.Replace("pt", "") + "(||)" + extractFontStyle + "(||)" + replaceSpace;
                                }
                            }

                        }


                        if (!nodes.ParentNode.Id.CaseInsensitiveContains("page"))
                        {

                            //get parent node
                            var nodeP = document.DocumentNode.SelectSingleNode("//div[@id='" + nodes.ParentNode.Id + "']");
                            //var nodeP = document.DocumentNode.SelectSingleNode("//div[@id='group002']");

                            var parserP = new Parser();
                            var stylesheetP = new StyleSheet();
                            //errorCheckNode = nodeP.Id + " hello " + nodeP.Attributes["style"].Value;

                            //errorCheckFile = nodeP.Id + " yolo " + nodeP.ParentNode.Id + nodeP.InnerText + nodeP.XPath;








                            stylesheetP = parser.Parse(".someClass{" + nodes.ParentNode.Attributes["style"].Value);

                            //stylesheetP = parserP.Parse(".someClass{" + nodeP.Attributes["style"].Value);



                            posiX_parent = stylesheetP.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                 .SelectMany(r => r.Declarations)
                                                                                 .FirstOrDefault(d => d.Name.CaseInsensitiveContains("left"))
                                                                                 .Term
                                                                                 .ToString();
                            posiY_parent = stylesheetP.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                .SelectMany(r => r.Declarations)
                                                                                .FirstOrDefault(d => d.Name.CaseInsensitiveContains("top"))
                                                                                .Term
                                                                                .ToString();
                            sizeX_parent = stylesheetP.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                             .SelectMany(r => r.Declarations)
                                                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("width"))
                                                             .Term
                                                             .ToString();
                            sizeY_parent = stylesheetP.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                             .SelectMany(r => r.Declarations)
                                                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("height"))
                                                             .Term
                                                             .ToString();


                            //if parent also contains % it has to go to its parent's parent until you dont get %
                            if (posiX_parent.Contains("%"))
                            {
                                int countP = 0;
                                //search the counter index for that parent. even if that parent has % ...if it had parent without %, posix will be calculated already
                                foreach (var nodesP in node.Descendants())
                                {

                                    if (nodesP.NodeType == HtmlNodeType.Element && nodesP.HasAttributes && nodesP.Attributes["id"] != null)
                                    {
                                        if (nodesP.Id.Equals(nodes.ParentNode.Id))
                                        {
                                            parentIndex = countP;
                                        }
                                        countP++;

                                    }

                                }


                                int parentX = xPositionFrmXl[parentIndex];
                                int parentY = yPositionFrmXl[parentIndex];
                                int parentWidth = xSizeFrmXl[parentIndex];
                                int parentHeight = ySizeFrmXl[parentIndex];

                                //MessageBox.Show(nodes.Id + " " + extractWidth + "width " + parentWidth + "parentIndex: "+ parentIndex + "parent name: " + nameFrmXl[parentIndex]);

                                double percentageX = parentWidth * Convert.ToDouble(extractPosiX.Replace("%", "")) / 100;
                                double percentageY = parentHeight * Convert.ToDouble(extractPosiY.Replace("%", "")) / 100;
                                double percentageW = parentWidth * Convert.ToDouble(extractWidth.Replace("%", "")) / 100;
                                double percentageH = parentHeight * Convert.ToDouble(extractHeight.Replace("%", "")) / 100;
                                //MessageBox.Show(percentageW.ToString());

                                xPositionFrmXl[counter] = parentX + (int)percentageX;
                                yPositionFrmXl[counter] = parentY + (int)percentageY;
                                xSizeFrmXl[counter] = (int)percentageW;
                                ySizeFrmXl[counter] = (int)percentageH;
                            }
                            else
                            {

                                if (extractPosiX.Contains("%"))
                                {
                                    int parentWidth = Convert.ToInt32(sizeX_parent.Replace("px", ""));
                                    int parentHeight = Convert.ToInt32(sizeY_parent.Replace("px", ""));

                                    int parentX = Convert.ToInt32(posiX_parent.Replace("px", ""));
                                    int parentY = Convert.ToInt32(posiY_parent.Replace("px", ""));

                                    //MessageBox.Show(extractWidth + "width " + parentWidth);
                                    double percentageX = parentWidth * Convert.ToDouble(extractPosiX.Replace("%", "")) / 100;
                                    double percentageY = parentHeight * Convert.ToDouble(extractPosiY.Replace("%", "")) / 100;
                                    double percentageW = parentWidth * Convert.ToDouble(extractWidth.Replace("%", "")) / 100;
                                    double percentageH = parentHeight * Convert.ToDouble(extractHeight.Replace("%", "")) / 100;
                                    //MessageBox.Show(percentageW.ToString());

                                    xPositionFrmXl[counter] = parentX + (int)percentageX;
                                    yPositionFrmXl[counter] = parentY + (int)percentageY;
                                    xSizeFrmXl[counter] = (int)percentageW;
                                    ySizeFrmXl[counter] = (int)percentageH;

                                    //MessageBox.Show(nodes.Id + " width " + xSizeFrmXl[counter] + " counter:" + counter);

                                }
                            }
                        }

                        if (extractPosiX.Contains("px"))
                        {
                            xPositionFrmXl[counter] = Convert.ToInt32(extractPosiX.Replace("px", ""));
                            yPositionFrmXl[counter] = Convert.ToInt32(extractPosiY.Replace("px", ""));
                            xSizeFrmXl[counter] = Convert.ToInt32(extractWidth.Replace("px", ""));
                            ySizeFrmXl[counter] = Convert.ToInt32(extractHeight.Replace("px", ""));
                        }



                        // forcolor
                        if (extractForeColor.Contains("#"))
                        {
                            foreColorFrmXl[counter] = extractForeColor;
                        }
                        else
                        {
                            if (extractForeColor.CaseInsensitiveContains("false") || extractForeColor.CaseInsensitiveContains("transparent"))
                            {
                                foreColorFrmXl[counter] = "";
                            }
                        }
                        //backcolor
                        if (extractBackColor.Contains("#"))
                        {
                            backColorFrmXl[counter] = extractBackColor;
                        }
                        else
                        {
                            if (extractBackColor.CaseInsensitiveContains("false") || extractBackColor.CaseInsensitiveContains("transparent"))
                            {
                                backColorFrmXl[counter] = "";
                            }
                        }

                        
                        
                        

                        //MessageBox.Show(counter + nodes.Id);
                        if (nodes.Id.CaseInsensitiveContains("rect") && !nodes.Id.CaseInsensitiveContains("round"))
                        {
                            typeFrmXl[counter] = "rectangle";

                            if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                             .SelectMany(r => r.Declarations)
                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("color")) != null)
                            {
                                foreColorFrmXl[counter] = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("color"))
                              .Term
                              .ToString();
                            }
                            if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                             .SelectMany(r => r.Declarations)
                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor")) != null)
                            {
                                backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                              .SelectMany(r => r.Declarations)
                                                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor"))
                                                              .Term
                                                              .ToString();

                                if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                             .SelectMany(r => r.Declarations)
                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("gradientfillcolor")) != null)
                                {
                                    //backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                    //                                                              .SelectMany(r => r.Declarations)
                                    //                                                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("gradientfillcolor"))
                                    //                                                              .Term
                                    //                                                              .ToString();
                                    configFrmXl[counter] = "gradient";
                                }
                            }

                            if (backColorFrmXl[counter].Equals("False"))
                            {
                                if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                             .SelectMany(r => r.Declarations)
                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("gradientfillcolor")) != null)
                                {
                                    backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                                  .SelectMany(r => r.Declarations)
                                                                                                  .FirstOrDefault(d => d.Name.CaseInsensitiveContains("gradientfillcolor"))
                                                                                                  .Term
                                                                                                  .ToString();
                                    configFrmXl[counter] = "gradient";
                                }
                                else
                                {
                                    //backColorFrmXl[counter] = "black";

                                }

                            }

                            if (backColorFrmXl[counter].Contains("-1"))
                            {
                                backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                                                                 .SelectMany(r => r.Declarations)
                                                                                                                                 .FirstOrDefault(d => d.Name.CaseInsensitiveContains("gradientfillcolor"))
                                                                                                                                 .Term
                                                                                                                                 .ToString();
                                configFrmXl[counter] = "gradient";
                            }
                            if (backColorFrmXl[counter].Equals(backgroundColorMD))
                            {
                                foreColorFrmXl[counter] = "black";
                            }

                            if (backColorFrmXl[counter].Contains("1.6777"))
                            {

                                backColorFrmXl[counter] = "#ffffff";
                            }
                            else if (backColorFrmXl[counter].Contains("1.26"))
                            {
                                backColorFrmXl[counter] = "#c0c0c0";

                            }

                            if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                             .SelectMany(r => r.Declarations)
                             .FirstOrDefault(d => d.Name.CaseInsensitiveContains("rotation")) != null)
                            {
                                rotation = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                              .SelectMany(r => r.Declarations)
                                                                                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("rotation"))
                                                                                              .Term
                                                                                              .ToString();
                                if(configFrmXl[counter] != null && !rotation.Equals("0"))
                                {
                                    configFrmXl[counter] = configFrmXl[counter] + "(||)rotation(||)" + rotation;
                                }
                                
                            }

                        }
                        else if (nodes.Id.CaseInsensitiveContains("oval"))
                        {
                            typeFrmXl[counter] = "circle";

                            if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                                 .SelectMany(r => r.Declarations)
                                                                                                 .FirstOrDefault(d => d.Name.CaseInsensitiveContains("color")) != null)
                                foreColorFrmXl[counter] = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                                     .SelectMany(r => r.Declarations)
                                                                                                     .FirstOrDefault(d => d.Name.CaseInsensitiveContains("color"))
                                                                                                     .Term
                                                                                                     .ToString();

                            if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                                 .SelectMany(r => r.Declarations)
                                                                                                 .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor")) != null)
                                backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                                     .SelectMany(r => r.Declarations)
                                                                                                     .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor"))
                                                                                                     .Term
                                                                                                     .ToString();

                            if (backColorFrmXl[counter].Contains("-1"))
                            {
                                backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                                                                 .SelectMany(r => r.Declarations)
                                                                                                                                 .FirstOrDefault(d => d.Name.CaseInsensitiveContains("gradientfillcolor"))
                                                                                                                                 .Term
                                                                                                                                 .ToString();
                                configFrmXl[counter] = "gradient";
                            }


                            if (backColorFrmXl[counter].Contains("1.6777"))
                            {

                                backColorFrmXl[counter] = "#ffffff";
                            }
                            else if (backColorFrmXl[counter].Contains("1.26"))
                            {
                                backColorFrmXl[counter] = "#c0c0c0";

                            }

                            if(backColorFrmXl[counter].Equals("False") || backColorFrmXl[counter].Equals(""))
                            {
                                //opposite of background color
                                backColorFrmXl[counter] = "black";

                            }
                        }
                        //i need to addd a button in extracted HMI
                        else if (nodes.Name.CaseInsensitiveContains("input") && nodes.Attributes["type"].Value.Equals("button"))
                        {
                            typeFrmXl[counter] = "input";

                            configFrmXl[counter] = "Arial(||)9(||)" + extractFontStyle + "(||)" + nodes.GetAttributeValue("value", "");
                            extractFontSize = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                .SelectMany(r => r.Declarations)
                                                                                .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-size"))
                                                                                .Term
                                                                                .ToString().Replace("pt", "");
                            extractFontStyle = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-weight"))
                                                                                   .Term
                                                                                   .ToString();

                            extractFont = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                            .SelectMany(r => r.Declarations)
                            .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-family"))
                            .Term
                            .ToString();

                            configFrmXl[counter] = extractFont + "(||)" + extractFontSize + "(||)" + extractFontStyle + "(||)" + nodes.GetAttributeValue("value", "");



                            //button color
                            if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("background-color")) != null)
                            {

                                backColorFrmXl[counter] = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.CaseInsensitiveContains("background-color"))
                               .Term
                               .ToString();
                            }


                            //text color
                            foreColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("textColor"))
                              .Term
                              .ToString();
                            if (foreColorFrmXl[counter].Equals("False"))
                            {
                                foreColorFrmXl[counter] = "black";
                            }
                        }
                        else if (nodes.Name.CaseInsensitiveContains("textarea") || nodes.Id.CaseInsensitiveContains("combo") || nodes.Id.CaseInsensitiveContains("cbo"))
                        {
                            typeFrmXl[counter] = "textbox";
                            if (nodes.InnerText == null)
                            {
                                configFrmXl[counter] = "Arial(||)9(||)regular(||)" + "";

                            }
                            else
                            {
                                extractFontSize = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-size"))
                                                                                   .Term
                                                                                   .ToString().Replace("pt", "");
                                extractFontStyle = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                                   .SelectMany(r => r.Declarations)
                                                                                   .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-weight"))
                                                                                   .Term
                                                                                   .ToString();

                                extractFont = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.CaseInsensitiveContains("font-family"))
                               .Term
                               .ToString();

                                configFrmXl[counter] = extractFont + "(||)" + extractFontSize + "(||)" + extractFontStyle + "(||)" + nodes.InnerText;

                            }

                            //textbox color
                            if(hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.Equals("fillColor")) != null)
                            {
                                backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                              .SelectMany(r => r.Declarations)
                                                              .FirstOrDefault(d => d.Name.Equals("fillColor"))
                                                              .Term
                                                              .ToString();
                            }
                            else
                            {
                                backColorFrmXl[counter] = "false";
                            }
                            

                            //text color
                            if(hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.Equals("textColor")) !=null)
                            {
                                foreColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.Equals("textColor"))
                              .Term
                              .ToString();
                            }
                            else if(stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.Equals("COLOR")) != null)
                            {
                                foreColorFrmXl[counter] = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.Equals("COLOR"))
                               .Term
                               .ToString();
                            }
                            else
                            {
                                foreColorFrmXl[counter] = "black";
                                //MessageBox.Show(nodes.Id);
                            }
                            
                            if (backColorFrmXl[counter].CaseInsensitiveContains("transparent"))
                            {
                                backColorFrmXl[counter] = "";
                            }
                            if (backColorFrmXl[counter].CaseInsensitiveContains("false"))
                            {
                                backColorFrmXl[counter] = "black";
                            }
                            if (foreColorFrmXl[counter].Equals("False"))
                            {
                                foreColorFrmXl[counter] = "black";
                            }
                        }
                        else if (nodes.Id.CaseInsensitiveContains("text") || nodes.Name.CaseInsensitiveContains("text") || nodes.Id.CaseInsensitiveContains("txt") || nodes.Id.Contains("question"))
                        {
                            typeFrmXl[counter] = "string";

                            //text color
                            if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.Equals("textColor")) != null)
                            {
                                backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.Equals("textColor"))
                              .Term
                              .ToString();
                            }
                            else if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.Equals("COLOR")) != null)
                            {
                                backColorFrmXl[counter] = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                               .SelectMany(r => r.Declarations)
                               .FirstOrDefault(d => d.Name.Equals("COLOR"))
                               .Term
                               .ToString();
                            }
                            else
                            {
                                backColorFrmXl[counter] = "black";
                                //MessageBox.Show(nodes.Id);
                            }


                        }
                        else if (nodes.Id.CaseInsensitiveContains("line"))
                        {

                            typeFrmXl[counter] = "line";

                            string linePoints = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("points"))
                              .Term
                              .ToString().TrimEnd();

                            //split by space and add the posix to even numbers and posiy to odd number of index
                            string[] pointsArray = linePoints.Split(' ');
                            float[] pointsA = Array.ConvertAll(pointsArray, s => float.Parse(s));

                            for (int i = 0; i < pointsArray.Length; i++)
                            {
                                if (i == 0)
                                {
                                    pointsA[i] = pointsA[i] * xSizeFrmXl[counter] / 100 + xPositionFrmXl[counter];
                                    //MessageBox.Show(xPositionFrmXl[counter].ToString());
                                }
                                //even
                                else if (i % 2 == 0)
                                {
                                    pointsA[i] = pointsA[i] * xSizeFrmXl[counter] / 100 + xPositionFrmXl[counter];
                                }
                                //odd
                                else
                                {
                                    pointsA[i] = pointsA[i] * ySizeFrmXl[counter] / 100 + yPositionFrmXl[counter];
                                }
                            }
                            // combine new strings
                            string combinedResult = string.Join(" ", pointsA);

                            //MessageBox.Show(points);

                            // MessageBox.Show(combinedResult);

                            configFrmXl[counter] = combinedResult + "(||)1";
                            //fill color and line color

                            
                            foreColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("linecolor"))
                              .Term
                              .ToString();

                            if (foreColorFrmXl[counter].Equals("False"))
                                {
                                foreColorFrmXl[counter] = "black";
                            }
                            //line width and color
                        }
                        else if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("points")) != null || nodes.Id.CaseInsensitiveContains("polygon"))
                        {
                            typeFrmXl[counter] = "polygon";
                            string points = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("points"))
                              .Term
                              .ToString().TrimEnd();

                            //split by space and add the posix to even numbers and posiy to odd number of index
                            string[] pointsArray = points.Split(' ');
                            float[] pointsA = Array.ConvertAll(pointsArray, s => float.Parse(s));

                            for (int i = 0; i < pointsArray.Length; i++)
                            {
                                if (i == 0)
                                {
                                    pointsA[i] = pointsA[i] * xSizeFrmXl[counter] / 100 + xPositionFrmXl[counter];
                                    //MessageBox.Show(xPositionFrmXl[counter].ToString());
                                }
                                //even
                                else if (i % 2 == 0)
                                {
                                    pointsA[i] = pointsA[i] * xSizeFrmXl[counter] / 100 + xPositionFrmXl[counter];
                                }
                                //odd
                                else
                                {
                                    pointsA[i] = pointsA[i] * ySizeFrmXl[counter] / 100 + yPositionFrmXl[counter];
                                }
                            }
                            // combine new strings
                            string combinedResult = string.Join(" ", pointsA);

                            //MessageBox.Show(points);

                            // MessageBox.Show(combinedResult);

                            configFrmXl[counter] = combinedResult + "(||)1";
                            //fill color and line color

                            backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor"))
                              .Term
                              .ToString();

                            if (backColorFrmXl[counter].Contains("1.6777"))
                            {

                                backColorFrmXl[counter] = "#ffffff";
                            }
                            else if (backColorFrmXl[counter].Contains("1.26"))
                            {
                                backColorFrmXl[counter] = "#c0c0c0";

                            }

                        }
                        //else if (nodes.Id.CaseInsensitiveContains("polygon"))
                        //{
                        //    typeFrmXl[counter] = "polygon";
                        //    string points = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                        //      .SelectMany(r => r.Declarations)
                        //      .FirstOrDefault(d => d.Name.CaseInsensitiveContains("points"))
                        //      .Term
                        //      .ToString();
                        //    configFrmXl[counter] = points + "(||)1";
                        //    //fill color and line color

                        //    backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                        //      .SelectMany(r => r.Declarations)
                        //      .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor"))
                        //      .Term
                        //      .ToString().Replace("#", "");
                        //}
                        else if (nodes.Name.CaseInsensitiveContains("img"))
                        {
                            string srcName = nodes.Attributes["src"].Value.Replace(".\\","");
                            //configFrmXl[counter] = srcName;
                            typeFrmXl[counter] = "image";
                            xPositionFrmXl[counter] = xPositionFrmXl[counter - 1];
                            yPositionFrmXl[counter] = yPositionFrmXl[counter - 1];
                            xSizeFrmXl[counter] = xSizeFrmXl[counter - 1];
                            ySizeFrmXl[counter] = ySizeFrmXl[counter - 1];

                            //copy image files to the save location
                            var dir = savePath + "extracted_images\\" + Path.GetDirectoryName(srcName);
                            if (!Directory.Exists(dir))
                            {
                                Directory.CreateDirectory(dir);
                            }

                            if (!File.Exists(savePath + "extracted_images\\" + srcName))
                            {
                                File.Copy(openFilePath + srcName, savePath + "extracted_images\\" + srcName);

                            }
                            configFrmXl[counter] = savePath + "extracted_images\\" + srcName;
                        }
                        else if (nodes.Id.CaseInsensitiveContains("picture"))
                        {
                            string srcName = nodes.Attributes["shapesrc"].Value.Replace(".\\", "");
                            typeFrmXl[counter] = "image";

                            //copy image files to the save location
                            var dir = savePath + "extracted_images\\" + Path.GetDirectoryName(srcName);
                            if (!Directory.Exists(dir))
                            {
                                Directory.CreateDirectory(dir);
                            }

                            if (!File.Exists(savePath + "extracted_images\\" + srcName))
                            {
                                File.Copy(openFilePath + srcName, savePath + "extracted_images\\" + srcName);

                            }

                            configFrmXl[counter] = savePath + "extracted_images\\" + srcName;

                        }
                        else if (nodes.ChildNodes.Count > 1)
                        {
                            if (nodes.ParentNode.Id.Equals("Page"))
                            {
                                mainGrpFlag = true;
                                typeFrmXl[counter] = "group";
                                mainGrpNodeId = nodes.Id;
                                mainGrpNodeParentId = nodes.ParentNode.Id;
                            }
                            else 
                            {
                                grpFlag = true;
                                typeFrmXl[counter] = "group";
                                grpNodeId = nodes.Id;
                                grpNodeParentId = nodes.ParentNode.Id;


                            }

                            foreColorFrmXl[counter] = "";
                            backColorFrmXl[counter] = "";

                           

                        }
                        else
                        {
                            typeFrmXl[counter] = "rectangle";
                            if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor")) != null)
                            {
                                backColorFrmXl[counter] = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                              .SelectMany(r => r.Declarations)
                                                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("fillcolor"))
                                                              .Term
                                                              .ToString();
                            }
                            else
                            {
                                if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                              .SelectMany(r => r.Declarations)
                                                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("color")) != null)
                                {
                                    backColorFrmXl[counter] = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                                  .SelectMany(r => r.Declarations)
                                                                  .FirstOrDefault(d => d.Name.CaseInsensitiveContains("color"))
                                                                  .Term
                                                                  .ToString();

                                }
                                else
                                {
                                    backColorFrmXl[counter] = "";
                                }
                            }

                                
                            

                            if (stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                              .SelectMany(r => r.Declarations)
                                                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("textcolor")) != null)
                            {
                                foreColorFrmXl[counter] = stylesheet.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                                                              .SelectMany(r => r.Declarations)
                                                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("textcolor"))
                                                              .Term
                                                              .ToString();
                            }
                            else
                            {
                                foreColorFrmXl[counter] = "";
                            }

                        }

                        nameFrmXl[counter] = nodes.Id;
                        //MessageBox.Show(nameFrmXl[counter]);

                        // get animation...                        
                       
                        //animation for binding
                        if (hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("hdxbindingid")) != null)
                        {
                            string bId = hdxProperty.StyleRules.Where(s => s.Selector.ToString() == ".someClass")
                              .SelectMany(r => r.Declarations)
                              .FirstOrDefault(d => d.Name.CaseInsensitiveContains("hdxbindingid"))
                              .Term
                              .ToString();

                            XmlDocument doc = new XmlDocument();
                            string str = openFilePath + fileNamePure + "_files\\";
                            doc.Load(str + "bindings.xml");
                            XmlElement root = doc.DocumentElement;
                            XmlNodeList xmlNodes = root.SelectNodes("binding");
                            string bindingID = "";
                            string previousBindingID = "";
                            string objectID = "";
                            foreach (XmlNode xmlNode in xmlNodes)
                            {
                                bindingID = xmlNode.Attributes["ID"].Value;
                                if (bindingID.Equals(bId))
                                {
                                    foreach (XmlNode xmlChildNode in xmlNode.ChildNodes)
                                    {
                                        //for nodes that contain more than 2 object iD
                                        int multipleObjectID = 0;

                                        if (xmlChildNode.Attributes["objectid"] != null)
                                        {
                                            multipleObjectID++;

                                            objectID = xmlChildNode.Attributes["objectid"].Value;


                                            //now we know the objectid we go through the datasource file
                                            XmlDocument dataFile = new XmlDocument();
                                            dataFile.Load(str + "DS_datasource1.dsd");
                                            XmlElement dataRoot = dataFile.DocumentElement;
                                            XmlNodeList dataNodes = dataRoot.SelectNodes("dataobject");

                                            int dataNodeCount = 0;
                                            foreach (XmlNode dataNode in dataNodes)
                                            {
                                                if (dataNode.Attributes["id"].Value.Equals(objectID))
                                                {

                                                    if (!previousBindingID.Equals(bindingID))
                                                    {
                                                        richTB.Text = dataNodes[dataNodeCount].InnerXml.Replace("</property>", "</property>\n");

                                                    }
                                                    else
                                                    {
                                                        richTB.AppendText(dataNodes[dataNodeCount].InnerXml.Replace("</property>", "</property>\n"));
                                                        //MessageBox.Show(previousBindingID + " vs" + bindingID + "== " +nodes.Id + " => " + dataNodes[dataNodeCount].InnerXml);

                                                    }
                                                    previousBindingID = bindingID;

                                                }

                                                dataNodeCount++;
                                            }
                                            //MessageBox.Show(filePath);
                                            var dir = savePath + "animation\\" + fileNamePure;
                                            if (!Directory.Exists(dir))
                                            {
                                                Directory.CreateDirectory(dir);
                                            }

                                            File.WriteAllText(Path.Combine(dir, nameFrmXl[counter] + ".rtf"), richTB.Rtf);
                                            animationFrmXl[counter] = fileNamePure + "\\" +nameFrmXl[counter] + ".rtf";
                                        }
                                    }

                                }

                            }


                        }
                        // if next node parent does not equal to the current node parent
                        //if(nodes.ChildNodes.Count > 1)
                        // {
                        //     counter++;
                        // }
                        // else
                        // {
                        //     counter += 2;
                        // }
                        //if (!(nodes.ParentNode.Id.Equals(mainGrpNodeId)) && counter > 1 && !typeFrmXl[counter - 1].Equals("group") && mainGrpFlag)
                        //{
                        //    typeFrmXl[counter - 1] = "*endgroup " + typeFrmXl[counter - 1];
                        //}
                        //main group ended
                        //endgroup

                        if (grpFlag)
                        {
                            // set * if maingrpflag is true and parent node is equal to the main group node
                            if (typeFrmXl[counter].Contains("group"))
                            {
                                typeFrmXl[counter] = "*" + typeFrmXl[counter];

                            }

                            if ((nodes.ParentNode.Id.Equals(grpNodeId)))
                            {
                                typeFrmXl[counter] = "*" + typeFrmXl[counter];
                            }                            

                        }

                        if (grpFlag && previousNode != null)
                        {

                            if (!nodes.Id.Equals(previousNode.ParentNode.Id) && counter > 1 && !nodes.ParentNode.Id.Equals(grpNodeId))
                            {

                                //mainGrpFlag = false;

                                if (!previousNode.ParentNode.Id.Equals(nodes.ParentNode.Id) && !typeFrmXl[counter].Contains("group"))
                                {

                                    //MessageBox.Show(previousNode.ParentNode.Id + " " + previousNode.Id + " " + nodes.Id);

                                    if (!typeFrmXl[counter - 1].Contains("endgroup"))
                                    {
                                        typeFrmXl[counter - 1] = "*endgroup " + typeFrmXl[counter - 1].Replace("*", "");

                                    }


                                    if (previousNode.ParentNode.Id.Equals(mainGrpNodeId))
                                    {
                                        grpFlag = false;
                                    }
                                }
                                if (nodes.ChildNodes.Count < 1 && nodes.ParentNode.Id.Equals(mainGrpNodeParentId))
                                {
                                    typeFrmXl[counter - 1] = "endgroup " + typeFrmXl[counter - 1];


                                    grpFlag = false;
                                }


                            }

                        }
                        //if ((nodes.ParentNode.Id.Equals(mainGrpNodeId)))
                        //{
                        //    typeFrmXl[counter] = "*" + typeFrmXl[counter];
                        //}

                        if (previousNode != null && mainGrpFlag)
                        {
                            
                            if (!nodes.Id.Equals(previousNode.ParentNode.Id) && !nodes.ParentNode.Id.Equals(mainGrpNodeId) && counter > 1)
                            {

                                //mainGrpFlag = false;
                                if (!previousNode.ParentNode.Id.Equals(nodes.ParentNode.Id) && !grpFlag && !typeFrmXl[counter].Contains("group"))
                                {

                                    //MessageBox.Show(previousNode.ParentNode.Id + " " + previousNode.Id + " " + nodes.Id);
                                    if(!typeFrmXl[counter-1].Contains("endgroup"))
                                    {
                                        typeFrmXl[counter - 1] = "endgroup " + typeFrmXl[counter - 1];

                                    }

                                    mainGrpFlag = false;
                                }

                                //if (nodes.ChildNodes.Count < 1 && nodes.ParentNode.Id.Equals("Page") && counter > 1)
                                //{
                                //    typeFrmXl[counter - 1] = "endgroup " + typeFrmXl[counter - 1];
                                //    mainGrpFlag = false;
                                //}

                            }

                        }

                       if (typeFrmXl[counter].Equals("group") && !previousNode.ParentNode.Id.Equals(mainGrpNodeParentId) && previousNode.ChildNodes.Count < 1 && counter > 1 )
                        {
                            if (!typeFrmXl[counter - 1].Contains("endgroup"))
                            {
                                typeFrmXl[counter - 1] = "endgroup " + typeFrmXl[counter - 1];

                            }

                            mainGrpFlag = false;
                        }

                        //if this node is final node
                        if (nodes.ParentNode.LastChild == nodes)
                        {

                            if (nodes.ParentNode.Id.Equals(mainGrpNodeId) && nodes.ChildNodes.Count < 1)
                            {
                                typeFrmXl[counter] = "endgroup " + typeFrmXl[counter];

                                


                            }
                            else if (nodes.ParentNode.Id.Equals(grpNodeId) && nodes.ChildNodes.Count < 1)
                            {
                                if(typeFrmXl[counter].Contains("*"))
                                {
                                    typeFrmXl[counter] = typeFrmXl[counter].Replace("*", "");
                                }
                                typeFrmXl[counter] = "*endgroup " + typeFrmXl[counter];


                            }
                        }

                        previousNode = nodes;


                        counter++;

                        // MessageBox.Show(nodes.Name + " " + nodes.Id + " " + extractPosiX + extractPosiX + extractWidth + extractFontSize + extractFont);
                    }
                }
            }
            catch(Exception error)
            {
                MessageBox.Show(error.ToString() + "\nError At node: " +errorCheckNode + "\nError file:" + errorCheckFile);
            }
            
        }

        private void saveTimer_Tick(object sender, EventArgs e)
        {
            statusLbl.Text = "Ready";
            saveTimer.Stop();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            while (listBox1.SelectedItems.Count > 0)
            {
                listBox1.Items.Remove(listBox1.SelectedItems[0]);
            }
        }

        static string PrettyXml(string xml)
        {
            var stringBuilder = new StringBuilder();

            var element = XElement.Parse(xml);

            var settings = new XmlWriterSettings();
            settings.OmitXmlDeclaration = true;
            settings.Indent = true;
            settings.NewLineOnAttributes = true;

            using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
            {
                element.Save(xmlWriter);
            }

            return stringBuilder.ToString();
        }

        //private void richTB_TextChanged(object sender, EventArgs e)
        //{
        //    CheckKeyword("property", Color.Red, 0);
        //    CheckKeyword("if", Color.Green, 0);
        //}



        //private void CheckKeyword(string word, Color color, int startIndex)
        //{
        //    if (richTB.Text.Contains(word))
        //    {
        //        int index = -1;
        //        int selectStart = richTB.SelectionStart;

        //        while ((index = richTB.Text.IndexOf(word, (index + 1))) != -1)
        //        {
        //            richTB.Select((index + startIndex), word.Length);
        //            richTB.SelectionColor = color;
        //            richTB.Select(selectStart, 0);
        //            richTB.SelectionColor = Color.Black;
        //        }
        //    }
        //}

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