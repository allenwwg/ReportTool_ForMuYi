//using ExcelAdapter = ExcelAdapter.ExcelAdapter;
using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Configuration;
namespace ReportGenerator
{
    public partial class MainWindow : Window
    {
        private const string OutputFile = "Report.csv";
        private const string AppConfigFile = "app.config";

        private readonly Dictionary<string, string> InputFiles = new Dictionary<string, string>()
        {
            { "Material", "material.csv" },
            { "Joint", "Joint.csv" },
            { "Adhesive-10", "Adhesive-10.csv" },
            { "Adhesive-16", "Adhesive-16.csv" },
            { "Adhesive-25", "Adhesive-25.csv" },
            { "Adhesive-50", "Adhesive-50.csv" },
            { "SpoolInfo", "SpoolInfo.csv" },
            { "Template", "template.xlsx" }
        };

        private readonly Dictionary<string, DataTable> Tables = new Dictionary<string, DataTable>()
        {
            { "Material", new DataTable() },
            { "Joint", new DataTable() },
            { "Adhesive-10", new DataTable() },
            { "Adhesive-16", new DataTable() },
            { "Adhesive-25", new DataTable() },
            { "Adhesive-50", new DataTable() },
            { "SpoolInfo", new DataTable() },
            { "Template", new DataTable() }
        };

        private readonly Dictionary<string, List<string>> TypeToShotcuts;

        private readonly List<string> productSers;

        public MainWindow()
        {
            //	Dictionary<string, string> strs = new Dictionary<string, string>()
            //	{
            //		{ "Material", "material.csv" },
            //		{ "Joint", "Joint.csv" },
            //		{ "Adhesive", "Adhesive.csv" },
            //		{ "SpoolInfo", "SpoolInfo.csv" },
            //		{ "Template", "template.xlsx" }
            //	};
            //	this.InputFiles = strs;
            //	Dictionary<string, DataTable> strs1 = new Dictionary<string, DataTable>()
            //	{
            //		{ "Material", new DataTable() },
            //		{ "Joint", new DataTable() },
            //		{ "Adhesive", new DataTable() },
            //		{ "Adhesive", new DataTable() },
            //		{ "SpoolInfo", new DataTable() },
            //		{ "Template", new DataTable() }
            //	};
            //	this.Tables = strs1;

            Dictionary<string, List<string>> strs2 = new Dictionary<string, List<string>>();
            List<string> strs3 = new List<string>()
            {
                "P1",
                "P2",
                "P3",
                "P4",
                "P5",
                "P6",
                "P7",
                "P8",
                "P9"
            };
            strs2.Add("Pipe", strs3);
            List<string> strs4 = new List<string>()
            {
                "F1",
                "F2",
                "F3",
                "FD",
                "FU",
                "FB",
                "FS",
                "FR"
            };
            strs2.Add("Flange", strs4);
            strs2.Add("Coupling", new List<string>()
            {
                "CP",
                "CD"
            });
            strs2.Add("Tee", new List<string>()
            {
                "TE",
                "TR"
            });
            strs2.Add("Reducer", new List<string>()
            {
                "RC",
                "RE"
            });
            strs2.Add("Lateral", new List<string>()
            {
                "LE",
                "LR"
            });
            strs2.Add("Cross", new List<string>()
            {
                "CE",
                "CR"
            });
            List<string> strs5 = new List<string>()
            {
                "E1",
                "E2",
                "E3",
                "E4",
                "E5",
                "E6",
                "E7",
                "E8",
                "E9"
            };
            strs2.Add("Elbow", strs5);
            strs2.Add("Exp. Coupling", new List<string>()
            {
                "DK",
                "DM",
                "EP"
            });
            List<string> strs6 = new List<string>()
            {
                "GS",
                "BS",
                "RS",
                "NS"
            };
            strs2.Add("Saddle", strs6);
            strs2.Add("Bell Mouth", new List<string>()
            {
                "BM"
            });
            strs2.Add("Miter Fitting", new List<string>()
            {
                "MITER"
            });
            strs2.Add("Attachment", new List<string>()
            {
                "NK",
                "OR",
                "GE"
            });
            this.TypeToShotcuts = strs2;
            this.productSers = new List<string>()
            {
                "MY-MOS16",
                "MY-MOS16C"
            };
            //base();
            this.InitializeComponent();

            txtPipe.Text = ConfigurationManager.AppSettings["Pipe"];
            txtAdh.Text = ConfigurationManager.AppSettings["Adh"];
            txtProject.Text = ConfigurationManager.AppSettings["Project"];
            txtPipeLossRate.Text = ConfigurationManager.AppSettings["Rate"];
        }

        private void AddAdhesiveAttachment(ExcelAdapter.ExcelAdapter ea, ref int rowIndexInMatSheet, List<List<string>> jointList, ref int iSeq)
        {
            IEnumerable<IGrouping<string, List<string>>> groupings =
                from w in jointList
                group w by string.Concat(w[7], w[4], w[1]); //fab, sereris, item no

            foreach (IGrouping<string, List<string>> group in groupings)
            {
                //a group has different size
                List<string> strs1 = group.Aggregate<List<string>>((List<string> a, List<string> b) =>
                {
                    double num = 0;
                    double num1 = 0;
                    string str = a[5].Trim();
                    string str1 = b[5].Trim();
                    double.TryParse(str, out num);
                    double.TryParse(str1, out num1);
                    a[5] = (num + num1).ToString();
                    return a;
                });
                iSeq++;
                strs1[0] = iSeq.ToString();
                if (!strs1[7].ToLower().Equals("loose"))
                {
                    //for FAB type
                    double kitCnt = 0;
                    if (double.TryParse(strs1[5], out kitCnt))
                    {
                        strs1[5] = Math.Ceiling(kitCnt).ToString();
                    }
                }
                else
                {
                    //for loose type
                    double kitCnt = 0;
                    if (double.TryParse(strs1[5], out kitCnt))
                    {
                        double adRate = 0;
                        if (double.TryParse(this.txtAdh.Text, out adRate))
                        {
                            kitCnt = (1 + adRate / 100) * kitCnt;
                            strs1[5] = Math.Ceiling(kitCnt).ToString();
                        }
                    }
                }
                List<string> dataList = new List<string>(strs1);
                dataList[2] = string.Empty;
                int num9 = rowIndexInMatSheet + 1;
                int num10 = num9;
                rowIndexInMatSheet = num9;
                ea.Add(1, num10, dataList);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            this.labIn.Content = folderBrowserDialog.SelectedPath;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            this.labOut.Content = folderBrowserDialog.SelectedPath;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(this.labIn.Content.ToString()))
            {
                System.Windows.Forms.MessageBox.Show("Invalid Input Directory");
            }
            else if (this.HasValidInputFiles())
            {
                try
                {
                    this.work();
                    System.Windows.Forms.MessageBox.Show("Done");
                }
                catch (Exception exception)
                {
                    System.Windows.Forms.MessageBox.Show(exception.Message);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Invalid Input Files");
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            base.Close();
        }

        private List<List<string>> convertTableToList(DataTable dt)
        {
            List<List<string>> lists = new List<List<string>>();

            if (dt == null) return lists;

            foreach (DataRow row in dt.Rows)
            {
                string str = row[0].ToString();
                char[] chrArray = new char[] { ',' };
                lists.Add(new List<string>(str.Split(chrArray)));
            }
            return lists;
        }

        private Task DoWorkAsync()
        {
            return Task.Run(new Action(this.test));
        }

        private string getItemTypeByNo(string itemNo)
        {
            string key = "Miter";
            string str = (itemNo.Length >= 6 ? itemNo.Substring(4, 2) : string.Empty);
            foreach (KeyValuePair<string, List<string>> typeToShotcut in this.TypeToShotcuts)
            {
                if (!typeToShotcut.Value.Contains(str))
                {
                    continue;
                }
                key = typeToShotcut.Key;
                break;
            }
            return key;
        }

        private bool HasValidInputFiles()
        {
            return true;
            //bool flag = false;
            //DirectoryInfo directoryInfo = new DirectoryInfo(this.labIn.Content.ToString());
            //IEnumerable<string> files =
            //    from f in (IEnumerable<FileInfo>)directoryInfo.GetFiles()
            //    select f.Name;
            //if (!this.InputFiles.Values.Any<string>((string f) => files.Contains<string>(f)))
            //{
            //    return flag;
            //}
            //flag = true;
            //return flag;
        }

        private void test()
        {
            Thread.Sleep(5000);
        }

        private void work()
        {
            char[] chrArray;
            FileDialog saveFileDialog = new SaveFileDialog()
            {
                CheckPathExists = true,
                Filter = "Excel 文件|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true
            };
            if (System.Windows.Forms.DialogResult.OK != saveFileDialog.ShowDialog())
            {
                return;
            }
            string outfileName = saveFileDialog.FileName;
            string inDir = this.labIn.Content.ToString();
            this.labOut.Content.ToString();
            using (ExcelAdapter.ExcelAdapter excelAdapter = new ExcelAdapter.ExcelAdapter())
            {
                foreach (KeyValuePair<string, string> inputFile in this.InputFiles)
                {
                    string inFilePath = Path.Combine(inDir, inputFile.Value);
                    if (!File.Exists(inFilePath))
                    {
                        string curDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                        inFilePath = Path.Combine(curDir, inputFile.Value);
                    }

                    if (!File.Exists(inFilePath)) continue;

                    bool flag = true;
                    if (inputFile.Key.Equals("Template"))
                    {
                        continue;
                    }
                    if (!inputFile.Key.Contains("Adhesive"))
                    {
                        excelAdapter.FirstRowHeader = false;
                    }
                    else
                    {
                        excelAdapter.FirstRowHeader = true;
                    }

                    excelAdapter.OpenFile(inFilePath, flag);
                    excelAdapter.Fill(this.Tables[inputFile.Key], inputFile.Key);
                }
            }
            string inTemplateFileName = Path.Combine(inDir, this.InputFiles["Template"]);
            if (!File.Exists(inTemplateFileName))
                inTemplateFileName = Path.Combine("./", this.InputFiles["Template"]);

            File.Copy(inTemplateFileName, outfileName, true);

            using (ExcelAdapter.ExcelAdapter excelAdapter1 = new ExcelAdapter.ExcelAdapter())
            {
                excelAdapter1.OpenFile(outfileName, false);
                int worksheetRowCount = excelAdapter1.GetWorksheetRowCount(1);
                DataTable item = this.Tables["Material"];
                List<List<string>> lists = new List<List<string>>();
                foreach (DataRow row in item.Rows)
                {
                    string str5 = row[0].ToString();
                    chrArray = new char[] { ',' };
                    string[] strArrays = str5.Split(chrArray);
                    List<string> strs = new List<string>(strArrays);
                    string itemTypeByNo = this.getItemTypeByNo(strArrays[1].ToString());
                    strs.Add(itemTypeByNo);
                    lists.Add(strs);
                }
                IEnumerable<IGrouping<string, List<string>>> groupings =
                    from l in lists
                    group l by string.Concat(l[l.Count - 1], l[2], l[6], l[1]);
                List<List<string>> lists1 = new List<List<string>>();
                foreach (IGrouping<string, List<string>> strs1 in groupings)
                {
                    lists1.Add(strs1.Aggregate<List<string>>((List<string> a, List<string> b) =>
                    {
                        double num = 0;
                        double num1 = 0;
                        string str = a[5].Replace("MM", "").Trim();
                        string str1 = b[5].Replace("MM", "").Trim();
                        double.TryParse(str, out num);
                        double.TryParse(str1, out num1);
                        a[5] = (num + num1).ToString();
                        return a;
                    }));
                }
                DataTable jointTable = this.Tables["Joint"];
                DataTable adhesiveTable10 = this.Tables["Adhesive-10"];
                DataTable adhesiveTable16 = this.Tables["Adhesive-16"];
                DataTable adhesiveTable25 = this.Tables["Adhesive-25"];
                DataTable adhesiveTable50 = this.Tables["Adhesive-50"];


                DataTable spoolInfoTable = this.Tables["SpoolInfo"];

                List<List<string>> adhesiveDataList10 = this.convertTableToList(adhesiveTable10);
                List<List<string>> adhesiveDataList16 = this.convertTableToList(adhesiveTable16);
                List<List<string>> adhesiveDataList25 = this.convertTableToList(adhesiveTable25);
                List<List<string>> adhesiveDataList50 = this.convertTableToList(adhesiveTable50);

                var adHesiveDataDic = new Dictionary<int, List<List<string>>>();
                adHesiveDataDic.Add(10, adhesiveDataList10);
                adHesiveDataDic.Add(16, adhesiveDataList16);
                adHesiveDataDic.Add(25, adhesiveDataList25);
                adHesiveDataDic.Add(50, adhesiveDataList50);

                //List<List<string>> adhesiveDataList = adhesiveDataList25.Concat(adhesiveDataList16).Concat(adhesiveDataList10).Concat(adhesiveDataList50).ToList();

                List<List<string>> spoolInfoDataList = this.convertTableToList(spoolInfoTable);
                List<List<string>> jointList = new List<List<string>>();
                foreach (DataRow dataRow in jointTable.Rows)
                {
                    string str6 = dataRow[0].ToString();
                    chrArray = new char[] { ',' };
                    string[] strArrays1 = str6.Split(chrArray);
                    List<string> matchedSpoolInfoData = null;
                    try
                    {
                        matchedSpoolInfoData = spoolInfoDataList.First<List<string>>((List<string> s) => s[0].Equals(strArrays1[0]));
                    }
                    catch (Exception ex)
                    {
                        string str7 = string.Format("Having problem to find series by spool No {0} in SpoolInfo.csv", strArrays1[0]);
                        System.Windows.Forms.MessageBox.Show(str7);
                        continue;
                    }
                    List<string> jointDataLine = new List<string>();
                    string spoolProductSer = matchedSpoolInfoData[4];
                    string size = strArrays1[2];
                    string jointType = strArrays1[3];
                    List<string> matchedAdhesiveData = null;
                    try
                    {
                        int sizeIndex = (size.Contains("\"") ? 1 : 0);
                        int jointTypeIndex = (jointType.ToLower().Equals("w") ? 2 : 3);
                        int productSerIndex = (spoolProductSer.Equals(this.productSers[0]) ? 4 : 5);
                        List<List<string>> adhesiveDataList = null;
                        if(spoolProductSer.Contains("10"))
                        {
                            adhesiveDataList = adHesiveDataDic[10];
                        }
                        else if(spoolProductSer.Contains("16"))
                        {
                            adhesiveDataList = adHesiveDataDic[16];
                        }
                        else if(spoolProductSer.Contains("25"))
                        {
                            adhesiveDataList = adHesiveDataDic[25];
                        }
                        else if(spoolProductSer.Contains("50"))
                        {
                            adhesiveDataList = adHesiveDataDic[50];
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("NO ad data found!");
                            continue;
                        }

                        matchedAdhesiveData = adhesiveDataList.First<List<string>>((List<string> a) => a[sizeIndex].Trim().Equals(size.Replace("\"", "").Trim()));
                        jointDataLine.Add("");
                        jointDataLine.Add(matchedAdhesiveData[productSerIndex]);
                        jointDataLine.Add(size.Replace("\"", "").Trim());
                        jointDataLine.Add((spoolProductSer.Equals("MY-MOS16C") ? "Conductive ADHESIVE" : "Non-Conductive ADHESIVE"));
                        jointDataLine.Add(spoolProductSer);
                        jointDataLine.Add(matchedAdhesiveData[jointTypeIndex]);
                        jointDataLine.Add("Kit");
                        jointDataLine.Add((jointType.ToLower().Equals("w") ? "FAB" : "LOOSE"));
                        jointList.Add(jointDataLine);
                    }
                    catch (Exception ex)
                    {
                        string str10 = string.Format("Having problem to find adhensive data by Size: {0}, JointType: {1}, Type: {2} in Adhensive.csv", size, jointType, spoolProductSer);
                        System.Windows.Forms.MessageBox.Show(str10);
                    }
                }
                IEnumerable<IGrouping<string, List<string>>> groupings1 =
                    from l in lists1
                    group l by l[l.Count - 1];
                bool adhesiveAttachementSetted = false;
                foreach (IGrouping<string, List<string>> strs5 in groupings1)
                {
                    string key = strs5.Key;
                    int num5 = worksheetRowCount + 1;
                    worksheetRowCount = num5;
                    excelAdapter1.Add(1, num5, new List<string>()
                    {
                        key
                    });
                    int num6 = 0;
                    foreach (List<string> item3 in strs5)
                    {
                        int num7 = num6 + 1;
                        num6 = num7;
                        item3[0] = num7.ToString();
                        item3[7] = item3[6];
                        if (!key.Equals("Pipe"))
                        {
                            item3[6] = "EA";
                        }
                        else
                        {
                            item3[6] = "MT";
                            double num8 = 0;
                            if (double.TryParse(item3[5], out num8))
                            {
                                double num9 = 0;
                                if (double.TryParse(this.txtPipe.Text, out num9))
                                {
                                    num8 = (1 + num9 / 100) * num8 / 1000;
                                    double num10 = Math.Round(num8, 2);
                                    item3[5] = num10.ToString();
                                }
                            }
                        }
                        if (item3[7].ToLower().Equals("erec"))
                        {
                            item3[7] = "LOOSE";
                        }
                        if (item3[1].Trim().Equals("0000GE"))
                        {
                            item3[6] = "G";
                        }
                        int num11 = worksheetRowCount + 1;
                        worksheetRowCount = num11;
                        excelAdapter1.Add(1, num11, item3);
                    }
                    if (!key.Equals("Attachment"))
                    {
                        continue;
                    }
                    this.AddAdhesiveAttachment(excelAdapter1, ref worksheetRowCount, jointList, ref num6);
                    adhesiveAttachementSetted = true;
                }

                //write attachment info
                if (!adhesiveAttachementSetted && jointList.Count > 0)
                {
                    int num12 = worksheetRowCount + 1;
                    worksheetRowCount = num12;
                    excelAdapter1.Add(1, num12, new List<string>()
                    {
                        "Attachment"
                    });
                    int iSeq = 0;
                    this.AddAdhesiveAttachment(excelAdapter1, ref worksheetRowCount, jointList, ref iSeq);
                    adhesiveAttachementSetted = true;
                }

                //write joint info
                int num14 = worksheetRowCount + 1;
                worksheetRowCount = num14;
                excelAdapter1.Add(1, num14, new List<string>()
                {
                    "Joint"
                });
                IEnumerable<IGrouping<string, List<string>>> groupings2 =
                    from w in jointList
                    group w by string.Concat(new string[] { w[7], "|", w[4], "|", w[2] }); //fab, sereris, size
                int jointSeq = 0;
                foreach (IGrouping<string, List<string>> group in groupings2)
                {
                    string key1 = group.Key;
                    chrArray = new char[] { '|' };
                    string[] strArrays2 = key1.Split(chrArray);
                    jointSeq++;
                    string empty = string.Empty;
                    string str11 = strArrays2[2];
                    string strJointType = "Joint";
                    string str13 = strArrays2[1];
                    string strCnt = group.Count<List<string>>().ToString();
                    string str15 = "EA";
                    string str16 = strArrays2[0];
                    int rowIndex = worksheetRowCount + 1;
                    worksheetRowCount = rowIndex;
                    List<string> rowData = new List<string>()
                    {
                        jointSeq.ToString(),
                        empty,
                        str11,
                        strJointType,
                        str13,
                        strCnt,
                        str15,
                        str16
                    };
                    excelAdapter1.Add(1, rowIndex, rowData);
                }

                //Write Spool info tab
                int worksheetRowCount1 = excelAdapter1.GetWorksheetRowCount(2);
                DataTable dataTable2 = this.Tables["SpoolInfo"];
                int num17 = 0;
                foreach (DataRow row1 in dataTable2.Rows)
                {
                    string str17 = row1[0].ToString();
                    chrArray = new char[] { ',' };
                    List<string> cellDataList = new List<string>(str17.Split(chrArray));
                    int num18 = num17 + 1;
                    num17 = num18;
                    cellDataList.Insert(0, num18.ToString());
                    cellDataList.Add("");
                    int num19 = worksheetRowCount1 + 1;
                    worksheetRowCount1 = num19;
                    excelAdapter1.Add(2, num19, cellDataList);
                }

                //Write Details tab
                int detailsSheetRowCnt = excelAdapter1.GetWorksheetRowCount(3);
                DataTable materialTable = this.Tables["Material"];
                //int detailsCellIndex = 0;
                foreach (DataRow r in materialTable.Rows)
                {
                    try
                    {
                        string strRowData = r[0].ToString();
                        chrArray = new char[] { ',' };
                        List<string> cellDataList = new List<string>(strRowData.Split(chrArray));
                        //int num18 = detailsCellIndex + 1;
                        //detailsCellIndex = num18;
                        //cellDataList.Insert(0, num18.ToString());
                        string strProj = txtProject.Text;
                        string strConcat = "/";
                        string strConcatedValue = strProj + strConcat + cellDataList.First();
                        string strCnt = cellDataList[cellDataList.Count - 2].ToString();
                        cellDataList.Insert(0, strConcatedValue);
                        cellDataList.Insert(0, strConcat);
                        cellDataList.Insert(0, strProj);
                        if (strCnt.Contains("MM"))
                            cellDataList.Insert(cellDataList.Count - 1, txtPipeLossRate.Text);
                        else
                            cellDataList.Insert(cellDataList.Count - 1, "0");

                        detailsSheetRowCnt++;
                        excelAdapter1.Add(3, detailsSheetRowCnt, cellDataList);
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }

                }

                excelAdapter1.Save();
            }
        }

        private void txtPipe_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}