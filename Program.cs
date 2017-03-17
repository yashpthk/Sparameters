using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Numerics;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace CapacitanceExtractor
{
    class Program
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            //Initializations
            long freq = 0;
            long freq2 = 0;
            int count = 0;
            int tableCount = 0;
            DataSet ds = new DataSet();
            string dynamicFileName = String.Empty;
            string[] freqPoints = new string[10];
            bool isComparisionMode = false;
            int numberOfMeasurementSets = 0;
            bool isAnotherRun = false;
            
            while (!operationSelector(ref freq, ref freq2, ref dynamicFileName, ref isComparisionMode, ref numberOfMeasurementSets, isAnotherRun))
            {
                #region B1505 Interpolation
                iB1505 B1505Data = new B1505();
                double[][] cValues_B1505 = B1505Data.getB1505InterpolatedData();
                #endregion

                if (isComparisionMode)
                {
                    if (isAnotherRun == false)
                    {
                        while (numberOfMeasurementSets > 0)
                        {
                            count = 0;
                            tableCount++;
                            BuildCapacitanceTable(ref ds, tableCount, freq, freq2, dynamicFileName + tableCount, isComparisionMode);
                            freqPoints[tableCount] = dynamicFileName + tableCount.ToString();

                            #region Add B1505 data and generate output file and Graphs
                            count = build1505DataAndGenerateOutputFile(count, tableCount, ds, dynamicFileName, cValues_B1505);
                            generateChartsAndPlots(ds.Tables[tableCount.ToString()], "Vds", "Cgd", tableCount + "_" + dynamicFileName);
                            Console.WriteLine("\n" + "Conversion Complete. Capacitances_" + tableCount + "_" + dynamicFileName + "MHz.txt file generated. \n");
                            #endregion
                            numberOfMeasurementSets--;
                        }
                        break;
                    }
                    else
                    {
                        isComparisionMode = false;
                        isAnotherRun = true;
                        Console.WriteLine("Measurement Comparator cannot be paired with other options." +
                            "\n Select any other option");
                        continue;
                    }
                }
                else
                {
                    isAnotherRun = true;
                    count = 0;
                    tableCount++;
                    BuildCapacitanceTable(ref ds, tableCount, freq, freq2, String.Empty, isComparisionMode);
                    freqPoints[tableCount] = dynamicFileName + "MHz";
                    #region Add B1505 data and generate output file and Graphs
                    count = build1505DataAndGenerateOutputFile(count, tableCount, ds, dynamicFileName, cValues_B1505);
                    generateChartsAndPlots(ds.Tables[tableCount.ToString()], "Vds", "Cgd", tableCount + "_" + dynamicFileName);
                    Console.WriteLine("\n" + "Conversion Complete. Capacitances_" + tableCount + "_" + dynamicFileName + "MHz.txt file generated. \n");
                    #endregion
                }
            }

            exportToExcel(ds, freqPoints);

            #region Error Graphs
            if (ds.Tables.Count != 0 && ds.Tables.Count > 1)
            {
                DataTable errorTable = new DataTable("ErrorTable");
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    errorTable.Columns.Add("errorCgd_" + freqPoints[i + 1]);
                    errorTable.Columns.Add("errorCgs_" + freqPoints[i + 1]);
                    errorTable.Columns.Add("errorCds_" + freqPoints[i + 1]);
                    errorTable.Columns.Add("errorCrss_" + freqPoints[i + 1]);
                    errorTable.Columns.Add("errorCiss_" + freqPoints[i + 1]);
                    errorTable.Columns.Add("errorCoss_" + freqPoints[i + 1]);
                }
                errorTable.Columns.Add("Vds", typeof(Double));
                for (int i = 1; i <= ds.Tables[0].Rows.Count; i++)
                {
                    errorTable.Rows.Add(errorTable.NewRow());
                }
                int setCount = 0;
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                    {
                        errorTable.Rows[j][setCount] = ds.Tables[i].Rows[j]["Cgd_Error"];
                        errorTable.Rows[j][setCount + 1] = ds.Tables[i].Rows[j]["Cgs_Error"];
                        errorTable.Rows[j][setCount + 2] = ds.Tables[i].Rows[j]["Cds_Error"];
                        errorTable.Rows[j][setCount + 3] = ds.Tables[i].Rows[j]["Crss_Error"];
                        errorTable.Rows[j][setCount + 4] = ds.Tables[i].Rows[j]["Ciss_Error"];
                        errorTable.Rows[j][setCount + 5] = ds.Tables[i].Rows[j]["Coss_Error"];
                        errorTable.Rows[j]["Vds"] = ds.Tables[i].Rows[j]["Vds"];
                    }
                    setCount = setCount + 6;
                }
                using (StreamWriter sw = File.CreateText(Environment.CurrentDirectory + @"\Output\Error_" + tableCount + "_" + dynamicFileName + "Mhz.txt"))
                {
                    foreach (DataRow row in errorTable.Rows)
                    {
                        sw.WriteLine(row[0] + "\t" + row[1] + "\t" + row[2] + "\t" +
                           row[4] + "\t" + row[5] + "\t" + row[6]);
                    }
                }
                    generateErrorGraphs(errorTable);
            } 
            #endregion
        }

        /// <summary>
        /// Build1505s the data and generate output file.
        /// </summary>
        /// <param name="count">The count.</param>
        /// <param name="tableCount">The table count.</param>
        /// <param name="ds">The ds.</param>
        /// <param name="dynamicFileName">Name of the dynamic file.</param>
        /// <param name="cValues_B1505">The c values B1505.</param>
        /// <returns></returns>
        private static int build1505DataAndGenerateOutputFile(int count, int tableCount, DataSet ds, 
                                                                string dynamicFileName, double[][] cValues_B1505)
        {
            using (StreamWriter sw = File.CreateText(Environment.CurrentDirectory + @"\Output\Capacitances_" + tableCount + "_" + dynamicFileName + "Mhz.txt"))
            {
                sw.WriteLine("Cgd    Cgs    Cds    Crss    Ciss    Coss");
                foreach (DataRow row in ds.Tables[tableCount.ToString()].Rows)
                {
                    row["Cgd_B1505_interpolated"] = cValues_B1505[0][count];
                    row["Cgd_Error"] = Math.Abs(((Double)row["Cgd_B1505_interpolated"] - (Double)row["Cgd"]) / (Double)row["Cgd_B1505_interpolated"]) * 100;
                    row["Cgs_B1505_interpolated"] = cValues_B1505[1][count];
                    row["Cgs_Error"] = Math.Abs(((Double)row["Cgs_B1505_interpolated"] - (Double)row["Cgs"]) / (Double)row["Cgs_B1505_interpolated"]) * 100;
                    row["Cds_B1505_interpolated"] = cValues_B1505[2][count];
                    row["Cds_Error"] = Math.Abs(((Double)row["Cds_B1505_interpolated"] - (Double)row["Cds"]) / (Double)row["Cds_B1505_interpolated"]) * 100;
                    row["Crss_B1505_interpolated"] = cValues_B1505[3][count];
                    row["Crss_Error"] = Math.Abs(((Double)row["Crss_B1505_interpolated"] - (Double)row["Crss"]) / (Double)row["Crss_B1505_interpolated"]) * 100;
                    row["Ciss_B1505_interpolated"] = cValues_B1505[4][count];
                    row["Ciss_Error"] = Math.Abs(((Double)row["Ciss_B1505_interpolated"] - (Double)row["Ciss"]) / (Double)row["Ciss_B1505_interpolated"]) * 100;
                    row["Coss_B1505_interpolated"] = cValues_B1505[5][count];
                    row["Coss_Error"] = Math.Abs(((Double)row["Coss_B1505_interpolated"] - (Double)row["Coss"]) / (Double)row["Coss_B1505_interpolated"]) * 100;
                    sw.WriteLine(row["Cgd"] + "\t" + row["Cgs"] + "\t" + row["Cds"] + "\t" +
                        row["Crss"] + "\t" + row["Ciss"] + "\t" + row["Coss"]);
                    count++;
                }
            }

            return count;
        }

        /// <summary>
        /// Operations the selector.
        /// </summary>
        /// <param name="freq">The freq.</param>
        /// <param name="freq2">The freq2.</param>
        /// <param name="dynamicFileName">Name of the dynamic file.</param>
        /// <param name="isComparisionMode">if set to <c>true</c> [is comparision mode].</param>
        /// <returns></returns>
        private static bool operationSelector(ref long freq, ref long freq2, ref string dynamicFileName, 
                                              ref bool isComparisionMode, ref  int numberOfMeasurementSets, bool isAnotherRun)
        {
            string selector;
            Console.WriteLine("Select Mode: " +
                "\n 1. Auto (Extraction Frequency 20MHz). \n 2. Manual single frequency. \n 3. Manual two frequencies." +
                " \n 4. Measurement Set Comparision Mode. \n Any other Key to exit the program.");
            selector = Console.ReadLine();
            switch (selector)
            {
                case "1":
                    {
                        freq = 20000000;
                        freq2 = freq;
                        dynamicFileName = "20";
                        break;
                    }
                case "2":
                    {
                        freq = readUserInputNumber("Enter Extraction frequency in Hz (should be an Integer): ");
                        freq2 = freq;
                        dynamicFileName = (freq / 1000000).ToString();
                        break;
                    }
                case "3":
                    {
                        freq = readUserInputNumber("Enter Lower Extraction frequency in Hz (should be an Integer): ");
                        freq2 = readUserInputNumber("Enter Higher Extraction frequency in Hz (should be an Integer): ");
                        dynamicFileName = (freq / 1000000).ToString() + "-" + (freq2 / 1000000).ToString();
                        break;
                    }
                case "4":
                    {
                        if (isAnotherRun == false)
                        {
                            numberOfMeasurementSets = (Int32)readUserInputNumber("Enter total number of results to compare: ");
                            freq = readUserInputNumber("Enter Extraction frequency in Hz (should be an Integer). ");
                            freq2 = freq;
                            dynamicFileName = ("MeasurementSet");
                            isComparisionMode = true;
                        }
                        break;
                    }
                default:
                    {
                        return true;
                    }
            }

            return false;
        }

        /// <summary>
        /// Builds the capacitance table.
        /// </summary>
        /// <param name="ds">The ds.</param>
        /// <param name="tableCount">The table count.</param>
        /// <param name="freq">The freq.</param>
        /// <param name="freq2">The freq2.</param>
        /// <param name="dynamicPath">The dynamic path.</param>
        /// <param name="isComparisionMode">if set to <c>true</c> [is comparision mode].</param>
        private static void BuildCapacitanceTable(ref DataSet ds, int tableCount, Int64 freq, Int64 freq2, 
                                                    string folderName, bool isComparisionMode)
        {
            ds.Tables.Add(new DataTable(tableCount.ToString()));
            int count = 0;
            double[] Vds = { 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 20, 30, 40, 50, 60 };
            string path = Environment.CurrentDirectory;
              #if DEBUG
                path = path + @"\Data";
#endif
            if (isComparisionMode)
            {
                path = path + @"\" + folderName;
#if DEBUG
                path = Environment.CurrentDirectory + @"\Data\" + folderName;
#endif
            }

            #region Build Columns
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Freq"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Vds", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("S11"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("S21"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("S12"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("S22"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("deltaS"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Y11"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Y21"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Y12"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Y22"));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cgd", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cgs", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cds", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Crss", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Ciss", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Coss", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cgd_B1505_interpolated", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cgs_B1505_interpolated", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cds_B1505_interpolated", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Crss_B1505_interpolated", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Ciss_B1505_interpolated", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Coss_B1505_interpolated", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cgd_Error", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cgs_Error", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Cds_Error", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Crss_Error", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Ciss_Error", typeof(Double)));
            ds.Tables[tableCount.ToString()].Columns.Add(new DataColumn("Coss_Error", typeof(Double)));
            #endregion

            #region Read S2P files for given frequency
            for (int i = 1; i < 25; i++)
                if (i < 10)
                    ConvertToDataTable(path + @"\0V" + i.ToString() + ".S2P", freq, ds.Tables[tableCount.ToString()]);
                else if (i > 9 && i < 11)
                    ConvertToDataTable(path + @"\" + (i - 9).ToString() + "V0.S2P", freq, ds.Tables[tableCount.ToString()]);
                else if (i >= 11 && i < 20)
                    ConvertToDataTable(path + @"\" + (i - 9).ToString() + "V0.S2P", freq2, ds.Tables[tableCount.ToString()]);
                else if (i >= 20)
                {
                    ConvertToDataTable(path + @"\" + (i + (count * 9)).ToString() + "V0.S2P", freq2, ds.Tables[tableCount.ToString()]);
                    count++;
                }
            #endregion

            #region Calculate Capacitances
            int rowCount = 0;
            foreach (DataRow row in ds.Tables[tableCount.ToString()].Rows)
            {
                Complex S11 = new Complex(Convert.ToDouble(row["S11"].ToString().Split(',')[0]), Convert.ToDouble(row["S11"].ToString().Split(',')[1]));
                Complex S21 = new Complex(Convert.ToDouble(row["S21"].ToString().Split(',')[0]), Convert.ToDouble(row["S21"].ToString().Split(',')[1]));
                Complex S12 = new Complex(Convert.ToDouble(row["S12"].ToString().Split(',')[0]), Convert.ToDouble(row["S12"].ToString().Split(',')[1]));
                Complex S22 = new Complex(Convert.ToDouble(row["S22"].ToString().Split(',')[0]), Convert.ToDouble(row["S22"].ToString().Split(',')[1]));

                Complex deltaS = (Complex)(50 * (Complex.One + S11) * 50 * (Complex.One + S22) - (50 * S12 * 50 * S21));
                row["deltaS"] = deltaS.Real.ToString() + "," + deltaS.Imaginary.ToString();

                Complex Y11 = ((Complex.One - S11) * 50 * (Complex.One + S22) + (S12 * 50 * S21)) / deltaS;
                row["Y11"] = Y11.Real.ToString() + "," + Y11.Imaginary.ToString();
                Complex Y21 = (-2 * 50 * S21) / deltaS;
                row["Y21"] = Y21.Real.ToString() + "," + Y21.Imaginary.ToString();
                Complex Y12 = (-2 * 50 * S12) / deltaS;
                row["Y12"] = Y12.Real.ToString() + "," + Y12.Imaginary.ToString();
                Complex Y22 = ((Complex.One + S11) * 50 * (Complex.One - S22) + (S12 * 50 * S21)) / deltaS;
                row["Y22"] = Y22.Real.ToString() + "," + Y22.Imaginary.ToString();

                Complex Za = 1 / (Y11 + Y21);
                Complex Zb = 1 / (Y22 + Y21);
                Complex Zc = 1 / (-Y21);

                row["Cgd"] = Convert.ToDouble((-1*Math.Pow(10,12) / (2 * Math.PI * Convert.ToDouble(row[0]) * Zc.Imaginary)).ToString());
                row["Cgs"] = Convert.ToDouble((-1 * Math.Pow(10, 12) / (2 * Math.PI * Convert.ToDouble(row[0]) * Za.Imaginary)).ToString());
                row["Cds"] = Convert.ToDouble((-1 * Math.Pow(10, 12) / (2 * Math.PI * Convert.ToDouble(row[0]) * Zb.Imaginary)).ToString());
                row["Ciss"] = Convert.ToDouble(row["Cgs"]) + Convert.ToDouble(row["Cgd"]);
                row["Coss"] = Convert.ToDouble(row["Cds"]) + Convert.ToDouble(row["Cgd"]);
                row["Crss"] = Convert.ToDouble(row["Cgd"]);
                row["Vds"] = Vds[rowCount];
                rowCount++;

                Console.WriteLine(row[0] + "\t" + row[6].ToString() + "\t" + row[11].ToString() + "\t" + Zc);
            } 
            #endregion
        }

        /// <summary>
        /// Converts to data table.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="frequency">The frequency.</param>
        /// <param name="tbl">The table.</param>
        private static void ConvertToDataTable(string filePath, Int64 frequency, DataTable tbl)
        {
            string[] lines = System.IO.File.ReadAllLines(filePath);
            lines = lines.Skip(5).ToArray();
            foreach (string line in lines)
            {
                var cols = line.Split('\t');
                //Console.WriteLine(line);
                if(Convert.ToDouble(cols[0]) > frequency)
                {
                    DataRow dr = tbl.NewRow();
                    dr[0] = Convert.ToDouble(cols[0]);
                    dr["S11"] = cols[1] + "," + cols[2];
                    dr["S21"] = cols[3] + "," + cols[4];
                    dr["S12"] = cols[5] + "," + cols[6];
                    dr["S22"] = cols[7] + "," + cols[8];
                    tbl.Rows.Add(dr);
                    break;
                }
            }            
        }

        /// <summary>
        /// Generates the charts and plots.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="xMember">The x member.</param>
        /// <param name="yMember">The y member.</param>
        /// <param name="fileName">Name of the file.</param>
        private static void generateChartsAndPlots(DataTable table, string xMember, string yMember, string fileName)
        {
            Font font = new Font(FontFamily.GenericSansSerif, 12);

            Chart chartCombined = new Chart();
            chartCombined.DataSource = table;
            chartCombined.Width = 1920;
            chartCombined.Height = 1080;

            #region Legend Data & Styles
            Legend legend = new Legend("legend1");
            legend.IsDockedInsideChartArea = true;
            legend.DockedToChartArea = "ChartAreaParasitic"; //Defined at the bottom.
            legend.Docking = Docking.Bottom;
            legend.LegendStyle = LegendStyle.Table;
            legend.TableStyle = LegendTableStyle.Tall;
            legend.IsTextAutoFit = false;
            legend.BackColor = Color.Transparent;
            legend.BorderColor = Color.SlateGray;
            legend.Font = font;
            legend.ItemColumnSpacing = 20;
            chartCombined.Legends.Add(legend);
            Legend legendInt = new Legend("legend2");
            legendInt.IsDockedInsideChartArea = true;
            legendInt.DockedToChartArea = "ChartAreaIntrinsic"; //Defined at the bottom.
            legendInt.Docking = Docking.Bottom;
            legendInt.LegendStyle = LegendStyle.Table;
            legendInt.TableStyle = LegendTableStyle.Tall;
            legendInt.IsTextAutoFit = false;
            legendInt.BackColor = Color.Transparent;
            legendInt.BorderColor = Color.SlateGray;
            legendInt.Font = font;
            legendInt.ItemColumnSpacing = 20;
            chartCombined.Legends.Add(legendInt);
            Legend legendError1 = new Legend("legendError1");
            legendError1.IsDockedInsideChartArea = true;
            legendError1.DockedToChartArea = "ChartAreaError"; //Defined at the bottom.
            legendError1.Docking = Docking.Left;
            legendError1.LegendStyle = LegendStyle.Table;
            legendError1.TableStyle = LegendTableStyle.Tall;
            legendError1.IsTextAutoFit = false;
            legendError1.BackColor = Color.Transparent;
            legendError1.BorderColor = Color.SlateGray;
            legendError1.Font = font;
            legendError1.ItemColumnSpacing = 20;
            Legend legendError2 = new Legend("legendError2");
            legendError2.IsDockedInsideChartArea = true;
            legendError2.DockedToChartArea = "ChartAreaErrorInt"; //Defined at the bottom.
            legendError2.Docking = Docking.Left;
            legendError2.LegendStyle = LegendStyle.Table;
            legendError2.TableStyle = LegendTableStyle.Tall;
            legendError2.IsTextAutoFit = false;
            legendError2.BackColor = Color.Transparent;
            legendError2.BorderColor = Color.SlateGray;
            legendError2.Font = font;
            legendError2.ItemColumnSpacing = 20; 
            #endregion

            #region Series Cgd

            Series serieCgd = new Series();
            serieCgd.Name = "Cgd";
            serieCgd.Color = Color.Red;
            serieCgd.BorderColor = Color.FromArgb(164, 164, 164);
            serieCgd.ChartType = SeriesChartType.Line;
            serieCgd.MarkerStyle = MarkerStyle.Circle;
            serieCgd.MarkerColor = Color.Transparent;
            serieCgd.MarkerBorderWidth = 2;
            serieCgd.MarkerBorderColor = Color.Red;
            serieCgd.MarkerSize = 8;
            serieCgd.BorderDashStyle = ChartDashStyle.Solid;
            serieCgd.BorderWidth = 2;
            serieCgd.XValueMember = xMember;
            serieCgd.YValueMembers = yMember;
            chartCombined.Series.Add(serieCgd);
            #endregion

            #region Series Cgs
            Series serieCgs = new Series();
            serieCgs.Name = "Cgs";
            serieCgs.Color = Color.Blue;
            serieCgs.BorderColor = Color.FromArgb(164, 164, 164);
            serieCgs.ChartType = SeriesChartType.Line;
            serieCgs.MarkerStyle = MarkerStyle.Square;
            serieCgs.MarkerColor = Color.Transparent;
            serieCgs.MarkerBorderWidth = 2;
            serieCgs.MarkerBorderColor = Color.Blue;
            serieCgs.MarkerSize = 8;
            serieCgs.BorderWidth = 2;
            serieCgs.BorderDashStyle = ChartDashStyle.Solid;
            serieCgs.XValueMember = xMember; //"freq"
            serieCgs.YValueMembers = "Cgs"; //"S11"
            chartCombined.Series.Add(serieCgs);
            #endregion

            #region Series Cds
            Series serieCds = new Series();
            serieCds.Name = "Cds";
            serieCds.Color = Color.Green;
            serieCds.BorderColor = Color.FromArgb(164, 164, 164);
            serieCds.ChartType = SeriesChartType.Line;
            serieCds.MarkerStyle = MarkerStyle.Triangle;
            serieCds.MarkerColor = Color.Transparent;
            serieCds.MarkerBorderWidth = 2;
            serieCds.MarkerBorderColor = Color.Green;
            serieCds.MarkerSize = 8;
            serieCds.BorderWidth = 2;
            serieCds.BorderDashStyle = ChartDashStyle.Solid;
            serieCds.XValueMember = xMember; //"freq"
            serieCds.YValueMembers = "Cds"; //"S11"
            chartCombined.Series.Add(serieCds);
            #endregion

            #region Series Cgd_B1505

            Series serieCgd_B1505 = new Series();
            serieCgd_B1505.Name = "Cgd_B1505";
            serieCgd_B1505.Color = Color.Red;
            serieCgd_B1505.BorderColor = Color.FromArgb(164, 164, 164);
            serieCgd_B1505.ChartType = SeriesChartType.Line;
            serieCgd_B1505.MarkerStyle = MarkerStyle.Diamond;
            serieCgd_B1505.MarkerColor = Color.Transparent;
            serieCgd_B1505.MarkerBorderWidth = 2;
            serieCgd_B1505.MarkerBorderColor = Color.Red;
            serieCgd_B1505.MarkerSize = 10;
            serieCgd_B1505.BorderDashStyle = ChartDashStyle.Dot;
            serieCgd_B1505.BorderWidth = 2;
            serieCgd_B1505.XValueMember = xMember;
            serieCgd_B1505.YValueMembers = "Cgd_B1505_interpolated";
            chartCombined.Series.Add(serieCgd_B1505);
            #endregion            

            #region Series Cgs_B1505

            Series serieCgs_B1505 = new Series();
            serieCgs_B1505.Name = "Cgs_B1505";
            serieCgs_B1505.Color = Color.Blue;
            serieCgs_B1505.BorderColor = Color.FromArgb(164, 164, 164);
            serieCgs_B1505.ChartType = SeriesChartType.Line;
            serieCgs_B1505.MarkerStyle = MarkerStyle.Diamond;
            serieCgs_B1505.MarkerColor = Color.Transparent;
            serieCgs_B1505.MarkerBorderWidth = 2;
            serieCgs_B1505.MarkerBorderColor = Color.Blue;
            serieCgs_B1505.MarkerSize = 10;
            serieCgs_B1505.BorderDashStyle = ChartDashStyle.Dot;
            serieCgs_B1505.BorderWidth = 2;
            serieCgs_B1505.XValueMember = xMember;
            serieCgs_B1505.YValueMembers = "Cgs_B1505_interpolated";
            chartCombined.Series.Add(serieCgs_B1505);
            #endregion            

            #region Series Cds_B1505

            Series serieCds_B1505 = new Series();
            serieCds_B1505.Name = "Cds_B1505";
            serieCds_B1505.Color = Color.Green;
            serieCds_B1505.BorderColor = Color.FromArgb(164, 164, 164);
            serieCds_B1505.ChartType = SeriesChartType.Line;
            serieCds_B1505.MarkerStyle = MarkerStyle.Diamond;
            serieCds_B1505.MarkerColor = Color.Transparent;
            serieCds_B1505.MarkerBorderWidth = 2;
            serieCds_B1505.MarkerBorderColor = Color.Green;
            serieCds_B1505.MarkerSize = 10;
            serieCds_B1505.BorderDashStyle = ChartDashStyle.Dot;
            serieCds_B1505.BorderWidth = 2;
            serieCds_B1505.XValueMember = xMember;
            serieCds_B1505.YValueMembers = "Cds_B1505_interpolated";
            chartCombined.Series.Add(serieCds_B1505);
            #endregion            

            #region Create chart area
            ChartArea ca = new ChartArea();
            ca.Name = "ChartAreaParasitic";
            ca.BackColor = Color.White;
            ca.BorderColor = Color.FromArgb(26, 59, 105);
            ca.BorderWidth = 0;
            ca.BorderDashStyle = ChartDashStyle.Solid;
            ca.AxisX = new Axis();
            ca.AxisX.Title = "Vds";
            ca.AxisX.TitleFont = font;
            ca.AxisX.IsLogarithmic = true;
            ca.AxisX.MinorGrid.Enabled = true;
            ca.AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca.AxisX.MinorGrid.LineColor = Color.LightGray;
            ca.AxisX.MinorGrid.LineWidth = 1;
            ca.AxisX.MinorGrid.Interval = 1;
            ca.AxisY = new Axis();
            ca.AxisY.Title = "Capacitances in pF";
            ca.AxisY.TitleFont = font;
            ca.AxisY.IsLogarithmic = true;
            ca.AxisY.MinorGrid.Enabled = true;
            ca.AxisY.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca.AxisY.MinorGrid.LineWidth = 1;
            ca.AxisY.MinorGrid.Interval = 1;
            chartCombined.ChartAreas.Add(ca);
            #endregion

            #region Series Crss

            Series serieCrss = new Series();
            serieCrss.Name = "Crss";
            serieCrss.Legend = "legend2";
            serieCrss.Color = Color.Red;
            serieCrss.BorderColor = Color.FromArgb(164, 164, 164);
            serieCrss.ChartType = SeriesChartType.Line;
            serieCrss.MarkerStyle = MarkerStyle.Circle;
            serieCrss.MarkerColor = Color.Transparent;
            serieCrss.MarkerBorderWidth = 2;
            serieCrss.MarkerBorderColor = Color.Red;
            serieCrss.MarkerSize = 8;
            serieCrss.BorderDashStyle = ChartDashStyle.Solid;
            serieCrss.BorderWidth = 2;
            serieCrss.XValueMember = xMember;
            serieCrss.YValueMembers = "Crss";
            serieCrss.ChartArea = "ChartAreaIntrinsic";
            chartCombined.Series.Add(serieCrss);
            #endregion

            #region Series Ciss
            Series serieCiss = new Series();
            serieCiss.Name = "Ciss";
            serieCiss.Legend = "legend2";
            serieCiss.Color = Color.Blue;
            serieCiss.BorderColor = Color.FromArgb(164, 164, 164);
            serieCiss.ChartType = SeriesChartType.Line;
            serieCiss.MarkerStyle = MarkerStyle.Square;
            serieCiss.MarkerColor = Color.Transparent;
            serieCiss.MarkerBorderWidth = 2;
            serieCiss.MarkerBorderColor = Color.Blue;
            serieCiss.MarkerSize = 8;
            serieCiss.BorderWidth = 2;
            serieCiss.BorderDashStyle = ChartDashStyle.Solid;
            serieCiss.XValueMember = xMember; //"freq"
            serieCiss.YValueMembers = "Ciss"; //"S11"
            serieCiss.ChartArea = "ChartAreaIntrinsic";
            chartCombined.Series.Add(serieCiss);
            #endregion

            #region Series Coss
            Series serieCoss = new Series();
            serieCoss.Name = "Coss";
            serieCoss.Legend = "legend2";
            serieCoss.Color = Color.Green;
            serieCoss.BorderColor = Color.FromArgb(164, 164, 164);
            serieCoss.ChartType = SeriesChartType.Line;
            serieCoss.MarkerStyle = MarkerStyle.Triangle;
            serieCoss.MarkerColor = Color.Transparent;
            serieCoss.MarkerBorderWidth = 2;
            serieCoss.MarkerBorderColor = Color.Green;
            serieCoss.MarkerSize = 8;
            serieCoss.BorderWidth = 2;
            serieCoss.BorderDashStyle = ChartDashStyle.Solid;
            serieCoss.XValueMember = xMember; //"freq"
            serieCoss.YValueMembers = "Coss"; //"S11"
            serieCoss.ChartArea = "ChartAreaIntrinsic";
            chartCombined.Series.Add(serieCoss);
            #endregion

            #region Series Crss_B1505
            Series serieCrssB1505 = new Series();
            serieCrssB1505.Name = "Crss_B1505";
            serieCrssB1505.Legend = "legend2";
            serieCrssB1505.Color = Color.Red;
            serieCrssB1505.BorderColor = Color.FromArgb(164, 164, 164);
            serieCrssB1505.ChartType = SeriesChartType.Line;
            serieCrssB1505.MarkerStyle = MarkerStyle.Diamond;
            serieCrssB1505.MarkerColor = Color.Transparent;
            serieCrssB1505.MarkerBorderWidth = 2;
            serieCrssB1505.MarkerBorderColor = Color.Red;
            serieCrssB1505.MarkerSize = 8;
            serieCrssB1505.BorderWidth = 2;
            serieCrssB1505.BorderDashStyle = ChartDashStyle.Dot;
            serieCrssB1505.XValueMember = xMember; //"freq"
            serieCrssB1505.YValueMembers = "Crss_B1505_interpolated"; //"S11"
            serieCrssB1505.ChartArea = "ChartAreaIntrinsic";
            chartCombined.Series.Add(serieCrssB1505);
            #endregion

            #region Series Ciss_B1505
            Series serieCissB1505 = new Series();
            serieCissB1505.Name = "Ciss_B1505";
            serieCissB1505.Legend = "legend2";
            serieCissB1505.Color = Color.Blue;
            serieCissB1505.BorderColor = Color.FromArgb(164, 164, 164);
            serieCissB1505.ChartType = SeriesChartType.Line;
            serieCissB1505.MarkerStyle = MarkerStyle.Diamond;
            serieCissB1505.MarkerColor = Color.Transparent;
            serieCissB1505.MarkerBorderWidth = 2;
            serieCissB1505.MarkerBorderColor = Color.Blue;
            serieCissB1505.MarkerSize = 8;
            serieCissB1505.BorderWidth = 2;
            serieCissB1505.BorderDashStyle = ChartDashStyle.Dot;
            serieCissB1505.XValueMember = xMember; //"freq"
            serieCissB1505.YValueMembers = "Ciss_B1505_interpolated"; //"S11"
            serieCissB1505.ChartArea = "ChartAreaIntrinsic";
            chartCombined.Series.Add(serieCissB1505);
            #endregion

            #region Series Coss_B1505
            Series serieCossB1505 = new Series();
            serieCossB1505.Name = "Coss_B1505";
            serieCossB1505.Legend = "legend2";
            serieCossB1505.Color = Color.Green;
            serieCossB1505.BorderColor = Color.FromArgb(164, 164, 164);
            serieCossB1505.ChartType = SeriesChartType.Line;
            serieCossB1505.MarkerStyle = MarkerStyle.Diamond;
            serieCossB1505.MarkerColor = Color.Transparent;
            serieCossB1505.MarkerBorderWidth = 2;
            serieCossB1505.MarkerBorderColor = Color.Green;
            serieCossB1505.MarkerSize = 8;
            serieCossB1505.BorderWidth = 2;
            serieCossB1505.BorderDashStyle = ChartDashStyle.Dot;
            serieCossB1505.XValueMember = xMember; //"freq"
            serieCossB1505.YValueMembers = "Coss_B1505_interpolated"; //"S11"
            serieCossB1505.ChartArea = "ChartAreaIntrinsic";
            chartCombined.Series.Add(serieCossB1505);
            #endregion

            #region Create chart area
            ChartArea ca2 = new ChartArea();
            ca2.Name = "ChartAreaIntrinsic";
            ca2.BackColor = Color.White;
            ca2.BorderColor = Color.FromArgb(26, 59, 105);
            ca2.BorderWidth = 0;
            ca2.BorderDashStyle = ChartDashStyle.Solid;
            ca2.AxisX = new Axis();
            ca2.AxisX.Title = "Vds";
            ca2.AxisX.TitleFont = font;
            ca2.AxisX.IsLogarithmic = true;
            ca2.AxisX.MinorGrid.Enabled = true;
            ca2.AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca2.AxisX.MinorGrid.LineColor = Color.LightGray;
            ca2.AxisX.MinorGrid.LineWidth = 1;
            ca2.AxisX.MinorGrid.Interval = 1;
            ca2.AxisY = new Axis();
            ca2.AxisY.Title = "Capacitances in pF";
            ca2.AxisY.TitleFont = font;
            ca2.AxisY.IsLogarithmic = true;
            ca2.AxisY.MinorGrid.Enabled = true;
            ca2.AxisY.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca2.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca2.AxisY.MinorGrid.LineWidth = 1;
            ca2.AxisY.MinorGrid.Interval = 1;
            chartCombined.ChartAreas.Add(ca2);
            #endregion

            #region Series ErrorCgd

            Series serieErrorCgd = new Series();
            serieErrorCgd.Name = "Cgd_Error";
            serieErrorCgd.Legend = "legendError1";
            serieErrorCgd.Color = Color.Red;
            serieErrorCgd.BorderColor = Color.FromArgb(164, 164, 164);
            serieErrorCgd.ChartType = SeriesChartType.Line;
            serieErrorCgd.MarkerStyle = MarkerStyle.Cross;
            serieErrorCgd.MarkerColor = Color.Transparent;
            serieErrorCgd.MarkerBorderWidth = 2;
            serieErrorCgd.MarkerBorderColor = Color.Red;
            serieErrorCgd.MarkerSize = 8;
            serieErrorCgd.BorderDashStyle = ChartDashStyle.Dot;
            serieErrorCgd.BorderWidth = 2;
            serieErrorCgd.XValueMember = xMember;
            serieErrorCgd.YValueMembers = "Cgd_Error";
            serieErrorCgd.ChartArea = "ChartAreaError";
            #endregion

            #region Series ErrorCgs

            Series serieErrorCgs = new Series();
            serieErrorCgs.Name = "Cgs_Error";
            serieErrorCgd.Legend = "legendError1";
            serieErrorCgs.Color = Color.Blue;
            serieErrorCgs.BorderColor = Color.FromArgb(164, 164, 164);
            serieErrorCgs.ChartType = SeriesChartType.Line;
            serieErrorCgs.MarkerStyle = MarkerStyle.Cross;
            serieErrorCgs.MarkerColor = Color.Transparent;
            serieErrorCgs.MarkerBorderWidth = 2;
            serieErrorCgs.MarkerBorderColor = Color.Blue;
            serieErrorCgs.MarkerSize = 8;
            serieErrorCgs.BorderDashStyle = ChartDashStyle.DashDot;
            serieErrorCgs.BorderWidth = 2;
            serieErrorCgs.XValueMember = xMember;
            serieErrorCgs.YValueMembers = "Cgs_Error";
            serieErrorCgs.ChartArea = "ChartAreaError";
            #endregion

            #region Series ErrorCds

            Series serieErrorCds = new Series();
            serieErrorCds.Name = "Cds_Error";
            serieErrorCgd.Legend = "legendError1";
            serieErrorCds.Color = Color.Green;
            serieErrorCds.BorderColor = Color.FromArgb(164, 164, 164);
            serieErrorCds.ChartType = SeriesChartType.Line;
            serieErrorCds.MarkerStyle = MarkerStyle.Cross;
            serieErrorCds.MarkerColor = Color.Transparent;
            serieErrorCds.MarkerBorderWidth = 2;
            serieErrorCds.MarkerBorderColor = Color.Green;
            serieErrorCds.MarkerSize = 8;
            serieErrorCds.BorderDashStyle = ChartDashStyle.Dash;
            serieErrorCds.BorderWidth = 2;
            serieErrorCds.XValueMember = xMember;
            serieErrorCds.YValueMembers = "Cds_Error";
            serieErrorCds.ChartArea = "ChartAreaError";
            #endregion

            #region Create chart area
            ChartArea ca3 = new ChartArea();
            ca3.Name = "ChartAreaError";
            ca3.BackColor = Color.White;
            ca3.BorderColor = Color.FromArgb(26, 59, 105);
            ca3.BorderWidth = 0;
            ca3.BorderDashStyle = ChartDashStyle.Solid;
            ca3.AxisX = new Axis();
            ca3.AxisX.Title = "Vds";
            ca3.AxisX.TitleFont = font;
            ca3.AxisX.IsLogarithmic = true;
            ca3.AxisX.MinorGrid.Enabled = true;
            ca3.AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca3.AxisX.MinorGrid.LineColor = Color.LightGray;
            ca3.AxisX.MinorGrid.LineWidth = 1;
            ca3.AxisX.MinorGrid.Interval = 1;
            ca3.AxisY = new Axis();
            ca3.AxisY.Title = "Capacitance Error";
            ca3.AxisY.TitleFont = font;
            ca3.AxisY.IsLogarithmic = false;
            ca3.AxisY.MinorGrid.Enabled = true;
            ca3.AxisY.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca3.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca3.AxisY.MinorGrid.LineWidth = 1;
            ca3.AxisY.MinorGrid.Interval = 1;
            #endregion

            #region Series ErrorCrss

            Series serieErrorCrss = new Series();
            serieErrorCrss.Name = "Crss_Error";
            serieErrorCrss.Legend = "legendError2";
            serieErrorCrss.Color = Color.Red;
            serieErrorCrss.BorderColor = Color.FromArgb(164, 164, 164);
            serieErrorCrss.ChartType = SeriesChartType.Line;
            serieErrorCrss.MarkerStyle = MarkerStyle.Cross;
            serieErrorCrss.MarkerColor = Color.Transparent;
            serieErrorCrss.MarkerBorderWidth = 2;
            serieErrorCrss.MarkerBorderColor = Color.Red;
            serieErrorCrss.MarkerSize = 8;
            serieErrorCrss.BorderDashStyle = ChartDashStyle.Dot;
            serieErrorCrss.BorderWidth = 2;
            serieErrorCrss.XValueMember = xMember;
            serieErrorCrss.YValueMembers = "Crss_Error";
            serieErrorCrss.ChartArea = "ChartAreaErrorInt";
            #endregion

            #region Series ErrorCgs

            Series serieErrorCiss = new Series();
            serieErrorCiss.Name = "Ciss_Error";
            serieErrorCiss.Legend = "legendError2";
            serieErrorCiss.Color = Color.Blue;
            serieErrorCiss.BorderColor = Color.FromArgb(164, 164, 164);
            serieErrorCiss.ChartType = SeriesChartType.Line;
            serieErrorCiss.MarkerStyle = MarkerStyle.Cross;
            serieErrorCiss.MarkerColor = Color.Transparent;
            serieErrorCiss.MarkerBorderWidth = 2;
            serieErrorCiss.MarkerBorderColor = Color.Blue;
            serieErrorCiss.MarkerSize = 8;
            serieErrorCiss.BorderDashStyle = ChartDashStyle.DashDot;
            serieErrorCiss.BorderWidth = 2;
            serieErrorCiss.XValueMember = xMember;
            serieErrorCiss.YValueMembers = "Ciss_Error";
            serieErrorCiss.ChartArea = "ChartAreaErrorInt";
            #endregion

            #region Series ErrorCds

            Series serieErrorCoss = new Series();
            serieErrorCoss.Name = "Coss_Error";
            serieErrorCoss.Legend = "legendError2";
            serieErrorCoss.Color = Color.Green;
            serieErrorCoss.BorderColor = Color.FromArgb(164, 164, 164);
            serieErrorCoss.ChartType = SeriesChartType.Line;
            serieErrorCoss.MarkerStyle = MarkerStyle.Cross;
            serieErrorCoss.MarkerColor = Color.Transparent;
            serieErrorCoss.MarkerBorderWidth = 2;
            serieErrorCoss.MarkerBorderColor = Color.Green;
            serieErrorCoss.MarkerSize = 8;
            serieErrorCoss.BorderDashStyle = ChartDashStyle.Dash;
            serieErrorCoss.BorderWidth = 2;
            serieErrorCoss.XValueMember = xMember;
            serieErrorCoss.YValueMembers = "Coss_Error";
            serieErrorCoss.ChartArea = "ChartAreaErrorInt";
            #endregion

            #region Create chart area
            ChartArea ca4 = new ChartArea();
            ca4.Name = "ChartAreaErrorInt";
            ca4.BackColor = Color.White;
            ca4.BorderColor = Color.FromArgb(26, 59, 105);
            ca4.BorderWidth = 0;
            ca4.BorderDashStyle = ChartDashStyle.Solid;
            ca4.AxisX = new Axis();
            ca4.AxisX.Title = "Vds";
            ca4.AxisX.TitleFont = font;
            ca4.AxisX.IsLogarithmic = true;
            ca4.AxisX.MinorGrid.Enabled = true;
            ca4.AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca4.AxisX.MinorGrid.LineColor = Color.LightGray;
            ca4.AxisX.MinorGrid.LineWidth = 1;
            ca4.AxisX.MinorGrid.Interval = 1;
            ca4.AxisY = new Axis();
            ca4.AxisY.Title = "Capacitance Error";
            ca4.AxisY.TitleFont = font;
            ca4.AxisY.IsLogarithmic = false;
            ca4.AxisY.MinorGrid.Enabled = true;
            ca4.AxisY.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca4.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca4.AxisY.MinorGrid.LineWidth = 1;
            ca4.AxisY.MinorGrid.Interval = 1;
            #endregion

            //databind...
            chartCombined.DataBind();
            //save result...
            chartCombined.SaveImage(Environment.CurrentDirectory + @"\Output\Combined_Capacitances_" + fileName + "MHz.png", ChartImageFormat.Png);

            Chart chartParasitic = new Chart();
            chartParasitic.DataSource = table;
            chartParasitic.Width = 1920;
            chartParasitic.Height = 1080;
            chartParasitic.Legends.Add(legend);
            chartParasitic.Series.Add(serieCgd);
            chartParasitic.Series.Add(serieCgs);
            chartParasitic.Series.Add(serieCds);
            chartParasitic.Series.Add(serieCgd_B1505);
            chartParasitic.Series.Add(serieCgs_B1505);
            chartParasitic.Series.Add(serieCds_B1505);
            chartParasitic.ChartAreas.Add(ca);
            chartParasitic.DataBind();
            chartParasitic.SaveImage(Environment.CurrentDirectory + @"\Output\Cgd_Cgs_Cds_" + fileName + "MHz.png", ChartImageFormat.Png);

            Chart chartIntrinsic = new Chart();
            chartIntrinsic.DataSource = table;
            chartIntrinsic.Width = 1920;
            chartIntrinsic.Height = 1080;
            chartIntrinsic.Legends.Add(legendInt);
            chartIntrinsic.Series.Add(serieCrss);
            chartIntrinsic.Series.Add(serieCiss);
            chartIntrinsic.Series.Add(serieCoss);
            chartIntrinsic.Series.Add(serieCrssB1505);
            chartIntrinsic.Series.Add(serieCissB1505);
            chartIntrinsic.Series.Add(serieCossB1505);
            chartIntrinsic.ChartAreas.Add(ca2);
            chartIntrinsic.DataBind();
            chartIntrinsic.SaveImage(Environment.CurrentDirectory + @"\Output\Crss_Ciss_Coss_" + fileName + "MHz.png", ChartImageFormat.Png);

            Chart chartErrorParasitic = new Chart();
            chartErrorParasitic.DataSource = table;
            chartErrorParasitic.Width = 1920;
            chartErrorParasitic.Height = 1080;
            chartErrorParasitic.Series.Add(serieErrorCgd);
            chartErrorParasitic.Series.Add(serieErrorCgs);
            chartErrorParasitic.Series.Add(serieErrorCds);
            chartErrorParasitic.Legends.Add(legendError1);
            chartErrorParasitic.ChartAreas.Add(ca3);
            chartErrorParasitic.DataBind();
            chartErrorParasitic.SaveImage(Environment.CurrentDirectory + @"\Output\Error_Par" + fileName + "MHz.png", ChartImageFormat.Png);

            Chart chartErrorIntrinsitic = new Chart();
            chartErrorIntrinsitic.DataSource = table;
            chartErrorIntrinsitic.Width = 1920;
            chartErrorIntrinsitic.Height = 1080;
            chartErrorIntrinsitic.Series.Add(serieErrorCrss);
            chartErrorIntrinsitic.Series.Add(serieErrorCiss);
            chartErrorIntrinsitic.Series.Add(serieErrorCoss);
            chartErrorIntrinsitic.Legends.Add(legendError2);
            chartErrorIntrinsitic.ChartAreas.Add(ca4);
            chartErrorIntrinsitic.DataBind();
            chartErrorIntrinsitic.SaveImage(Environment.CurrentDirectory + @"\Output\Error_Int" + fileName + "MHz.png", ChartImageFormat.Png);
        }

        /// <summary>
        /// Generates the error graphs.
        /// </summary>
        /// <param name="errorTable">The error table.</param>
        private static void generateErrorGraphs(DataTable errorTable)
        {
            Font font = new Font(FontFamily.GenericSansSerif, 12);

            Chart errorGraph = new Chart();
            errorGraph.DataSource = errorTable;
            errorGraph.Width = 1920;
            errorGraph.Height = 1080;

            Legend legend = new Legend("legend");
            legend.IsDockedInsideChartArea = true;
            legend.DockedToChartArea = "ChartAreaError"; //Defined at the bottom.
            legend.Docking = Docking.Top;
            legend.LegendStyle = LegendStyle.Column;
            legend.IsTextAutoFit = false;
            legend.BackColor = Color.Transparent;
            legend.BorderColor = Color.SlateGray;
            legend.Font = font;
            legend.ItemColumnSpacing = 20;
            errorGraph.Legends.Add(legend);

            #region Series Error_Set

            Random rand = new Random();
            for (int columnCount = 0; columnCount < (errorTable.Columns.Count - 1); columnCount = columnCount + 6)
            {
                errorGraph.Series.Add(buildErrorSeries(errorTable.Columns[columnCount].ColumnName, ref rand, columnCount + 1));
                errorGraph.Series.Add(buildErrorSeries(errorTable.Columns[columnCount + 1].ColumnName, ref rand, columnCount + 1));
                errorGraph.Series.Add(buildErrorSeries(errorTable.Columns[columnCount + 2].ColumnName, ref rand, columnCount + 1));
            }

            #endregion

            #region Create chart area
            buildChartArea(font);
            errorGraph.ChartAreas.Add(buildChartArea(font));
            #endregion

            //databind...
            errorGraph.DataBind();
            //save result...
            errorGraph.SaveImage(Environment.CurrentDirectory + @"\Output\Combined_Error_Par.png", ChartImageFormat.Png);

            errorGraph = new Chart();
            errorGraph.DataSource = errorTable;
            errorGraph.Width = 1920;
            errorGraph.Height = 1080;
            Legend legend2 = new Legend("legend2");
            legend2.IsDockedInsideChartArea = true;
            legend2.DockedToChartArea = "ChartAreaError"; //Defined at the bottom.
            legend2.Docking = Docking.Left;
            legend2.LegendStyle = LegendStyle.Column;
            legend2.IsTextAutoFit = false;
            legend2.BackColor = Color.Transparent;
            legend2.BorderColor = Color.SlateGray;
            legend2.Font = font;
            legend2.ItemColumnSpacing = 20;
            errorGraph.Legends.Add(legend2);

            #region Series Error_Set

            rand = new Random();
            for (int columnCount = 0; columnCount < (errorTable.Columns.Count - 1); columnCount = columnCount + 6)
            {
                errorGraph.Series.Add(buildErrorSeries(errorTable.Columns[columnCount + 3].ColumnName, ref rand, columnCount + 1));
                errorGraph.Series.Add(buildErrorSeries(errorTable.Columns[columnCount + 4].ColumnName, ref rand, columnCount + 1));
                errorGraph.Series.Add(buildErrorSeries(errorTable.Columns[columnCount + 5].ColumnName, ref rand, columnCount + 1));
            }

            #endregion

            #region Create chart area
            buildChartArea(font);
            errorGraph.ChartAreas.Add(buildChartArea(font));
            #endregion
            
            //databind...
            errorGraph.DataBind();
            //save result...
            errorGraph.SaveImage(Environment.CurrentDirectory + @"\Output\Combined_Error_Int.png", ChartImageFormat.Png);
        }

        private static ChartArea buildChartArea(Font font)
        {
            ChartArea ca = new ChartArea();
            ca.Name = "ChartAreaError";
            ca.BackColor = Color.White;
            ca.BorderColor = Color.FromArgb(26, 59, 105);
            ca.BorderWidth = 0;
            ca.BorderDashStyle = ChartDashStyle.Solid;
            ca.AxisX = new Axis();
            ca.AxisX.Title = "Vds";
            ca.AxisX.TitleFont = font;
            ca.AxisX.IsLogarithmic = true;
            ca.AxisX.MinorGrid.Enabled = true;
            ca.AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca.AxisX.MinorGrid.LineColor = Color.LightGray;
            ca.AxisX.MinorGrid.LineWidth = 1;
            ca.AxisX.MinorGrid.Interval = 1;
            ca.AxisY = new Axis();
            ca.AxisY.Maximum = 100;
            ca.AxisY.Title = "Error";
            ca.AxisY.TitleFont = font;
            ca.AxisY.IsLogarithmic = false;
            ca.AxisY.MinorGrid.Enabled = true;
            ca.AxisY.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            ca.AxisY.MinorGrid.LineColor = Color.LightGray;
            ca.AxisY.MinorGrid.LineWidth = 1;
            ca.AxisY.MinorGrid.Interval = 1;
            return ca;
        }

        /// <summary>
        /// Builds the error series.
        /// </summary>
        /// <param name="columnCount">The column count.</param>
        /// <param name="rand">The rand.</param>
        /// <returns></returns>
        private static Series buildErrorSeries(string serieName, ref Random rand, int pass)
        {
            Series serieCgd = new Series();
            Color color = Color.FromArgb(rand.Next(128), rand.Next(128), rand.Next(128));
            serieCgd.Name = serieName;
            serieCgd.Color = color;
            serieCgd.BorderColor = Color.FromArgb(164, 164, 164);
            serieCgd.ChartType = SeriesChartType.Line;
            serieCgd.MarkerStyle = (MarkerStyle)((pass > 6) ? pass / 6 : pass);
            serieCgd.MarkerColor = Color.Transparent;
            serieCgd.MarkerBorderWidth = 2;
            serieCgd.MarkerBorderColor = color;
            serieCgd.MarkerSize = 8;
            serieCgd.BorderDashStyle = (ChartDashStyle)((pass == 1) ? 5 : ((pass > 5) ? pass/5 : pass));
            serieCgd.BorderWidth = 2;
            serieCgd.XValueMember = "Vds";
            serieCgd.YValueMembers = serieName;
            return serieCgd;
        }

        /// <summary>
        /// Exports to excel.
        /// </summary>
        /// <param name="ds">The ds.</param>
        /// <param name="freqPoints">The freq points.</param>
        private static void exportToExcel(DataSet ds, string[] freqPoints)
        {
            int tableCount = 1;
            Console.WriteLine("Now generating excel file. This might take a while. \n" +
                "Check file overwrite confirmation dialog box if file already exists at location");
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = freqPoints[tableCount];
                excelWorkSheet.Cells[1, 1] = freqPoints[tableCount]+ ", Generated By CapacitanceExtractor.Exe, " 
                    + "Authors - Yash Pathak, Cristino Salcines, Date;Time " + DateTime.Now.ToString("yyyy/MM/dd;HH:mm:ss");

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[2, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 3, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
                tableCount++;
            }

            excelWorkBook.SaveAs(Environment.CurrentDirectory + @"\Output\Capacitances.xlsx");
            excelWorkBook.Close();
            excelApp.Quit();
        }

        /// <summary>
        /// Reads the user input number.
        /// </summary>
        /// <returns></returns>
        private static long readUserInputNumber(string displayMessage)
        {
            bool isValid = false;
            long input = -1;
            while (isValid == false)
            {
                Console.WriteLine(displayMessage);
                string tempInput = String.Empty;
                tempInput = Console.ReadLine();
                input = (tempInput == String.Empty || tempInput == null) ? 0 : Convert.ToInt64(tempInput);
                isValid = (input > 0) ? true : false;
            }
            return input;
        }
    }
}
