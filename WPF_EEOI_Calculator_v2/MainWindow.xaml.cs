using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.Axes;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using MNet = MathNet.Numerics;

namespace WPF_EEOI_Calculator_v2
{
    public partial class MainWindow : Window
    {
        public Company company { get; set; } = new Company();
        public string FileName = "";

        public MainWindow()
        {
            InitializeComponent();
            company.Vessels.Add(new Vessel());
            lbVessels1.SelectedIndex = 0;
        }

        #region Main Menu
        private void mnuNew_Click(object sender, RoutedEventArgs e)
        {
            company = new Company();
            company.Vessels.Add(new Vessel());
            lbVessels1.ItemsSource = company.Vessels;
            lbVessels1.SelectedIndex = 0;
            this.DataContext = company;
            this.Title = "Tecnitas - EEOI Calculator v2.0";
        }

        private void mnuOpen_Click(object sender, RoutedEventArgs e)
        {
            Company tempCompany = new Company();
            tempCompany = FileOperation.OpenXMLObject<Company>();
            if (tempCompany != null)
            {
                company = tempCompany;
                this.DataContext = company;
                lbVessels1.SelectedIndex = company.Vessels.Count() - 1;
                this.Title = "Tecnitas - EEOI Calculator v2.0 -" + FileOperation.FileName;
            }
        }

        private void mnuAppend_Click(object sender, RoutedEventArgs e)
        {
            Company appendedCompany = new Company();
            appendedCompany = FileOperation.OpenXMLObject<Company>();
            if (appendedCompany != null)
            {
                foreach (Vessel av in appendedCompany.Vessels)
                {
                    if (!company.Vessels.Contains(av, new VesselIMOComparer()))
                    {
                        company.Vessels.Add(av);
                    }
                }

                foreach (Vessel av in appendedCompany.Vessels)
                {
                    if (company.Vessels.Contains(av, new VesselIMOComparer()))
                    {
                        Vessel currentVessel = company.Vessels.Where(x => x.IMO == av.IMO).First();
                        var voyages = currentVessel.Voyages.Union<Voyage>(av.Voyages).ToArray<Voyage>();

                        for (int i = 0; i < voyages.Count(); i++)
                        {
                            if (!company.Vessels.Where(x => x.IMO == av.IMO).First().Voyages.Contains(voyages[i]))
                                company.Vessels.Where(x => x.IMO == av.IMO).First().Voyages.Add(voyages[i]);
                        }
                    }
                }
            }
        }

        private void mnuSave_Click(object sender, RoutedEventArgs e)
        {
            FileOperation.SaveObjectToXML<Company>(company);
            this.Title = "EEOI Calculator v2.0 -" + FileOperation.FileName;
        }

        private void mnuSaveAs_Click(object sender, RoutedEventArgs e)
        {
            FileOperation.SaveAsObjectToXML<Company>(company);
            this.Title = "EEOI Calculator v2.0 -" + FileOperation.FileName;
        }


        private void mnuExit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void mnuHelp_Click(object sender, RoutedEventArgs e)
        {
            string exePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            string filename = "EEOI Calculator 2.0_Manual.pdf";
            string path = exePath + "\\Resources\\" + filename;
            System.Diagnostics.Process.Start(path);
        }

        private void mnuAbout_Click(object sender, RoutedEventArgs e)
        {
            About about = new About();
            about.Owner = this;
            about.ShowDialog();
        }

        private void mnuExportXLS_Click(object sender, RoutedEventArgs e)
        {
            Parallel.Invoke(
           () =>
           {
               Thread thread = new Thread(new ThreadStart(() => { CreateExcelFile(); }));
               thread.Start();
           });
        }
        #endregion

        #region Buttons
        private void btnAddVessel_Click(object sender, RoutedEventArgs e)
        {
            lbVessels1.SelectedIndex = -1;
            Vessel v = new Vessel();
            company.Vessels.Add(v);
            lbVessels1.SelectedItem = v;
        }

        private void btnRemoveVessel_Click(object sender, RoutedEventArgs e)
        {
            if (lbVessels1.SelectedIndex > -1)
                company.Vessels.RemoveAt(lbVessels1.SelectedIndex);
        }
        #endregion

        #region DGV0 Context Menu
        private void mnuGroupVesselType_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV0.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
                cv.GroupDescriptions.Add(new PropertyGroupDescription("VesselType"));
            }
        }

        private void mnuGroupVesselFlag_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV0.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
                cv.GroupDescriptions.Add(new PropertyGroupDescription("Flag"));
            }
        }

        private void mnuUnGroupDGV0_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV0.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
            }
        }
        #endregion

        #region DGV1 Context Menu
        private void mnuEnableAll_Click(object sender, RoutedEventArgs e)
        {
            if (DGV1.ItemsSource != null)
            {
                foreach (Voyage v in DGV1.ItemsSource)
                {
                    v.IsEnabled = true;
                    ((Vessel)DGV0.SelectedItem).NotifyChange("");
                }
            }
        }

        private void mnuDisableAll_Click(object sender, RoutedEventArgs e)
        {
            if (DGV1.ItemsSource != null)
            {
                foreach (Voyage v in DGV1.ItemsSource)
                {
                    v.IsEnabled = false;
                    ((Vessel)DGV0.SelectedItem).NotifyChange("");
                }
            }
        }

        private void mnuEnableAllSelected_Click(object sender, RoutedEventArgs e)
        {
            if (DGV1.ItemsSource != null)
            {
                foreach (var v in DGV1.SelectedItems)
                {
                    if (v is Voyage)
                    {
                        ((Voyage)v).IsEnabled = true;
                        ((Vessel)DGV0.SelectedItem).NotifyChange("");
                    }
                }
            }
        }

        private void mnuDisableAllSelected_Click(object sender, RoutedEventArgs e)
        {
            if (DGV1.ItemsSource != null)
            {
                foreach (var v in DGV1.SelectedItems)
                {
                    if (v is Voyage)
                    {
                        ((Voyage)v).IsEnabled = false;
                        ((Vessel)DGV0.SelectedItem).NotifyChange("");
                    }
                }
            }
        }

        private void mnuGroupId_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV1.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
                cv.GroupDescriptions.Add(new PropertyGroupDescription("Id"));
            }
        }

        private void mnuGroupPort_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV1.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
                cv.GroupDescriptions.Add(new PropertyGroupDescription("DeparturePort"));
            }
        }

        private void mnuGroupVoyageType_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV1.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
                cv.GroupDescriptions.Add(new PropertyGroupDescription("VoyageType"));
            }
        }

        private void mnuGroupDistance_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV1.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
                cv.GroupDescriptions.Add(new PropertyGroupDescription("Distance"));
            }
        }

        private void mnuGroupCargoMass_Click(object sender, RoutedEventArgs e)
        {
            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV1.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
                cv.GroupDescriptions.Add(new PropertyGroupDescription("CargoMass"));
            }
        }

        private void mnuUnGroup_Click(object sender, RoutedEventArgs e)
        {

            CollectionView cv = (CollectionView)CollectionViewSource.GetDefaultView(DGV1.ItemsSource);
            if (cv != null)
            {
                cv.GroupDescriptions.Clear();
            }
        }
        #endregion

        #region lbVessels2 Context Menu
        private void mnuVesselEnableAllSelected_Click(object sender, RoutedEventArgs e)
        {
            if (lbVessels2.SelectedItems != null)
            {
                foreach (Vessel vessel in lbVessels2.SelectedItems)
                {
                    vessel.IsSelected = true;
                }
            }
        }

        private void mnuVesselDisableAllSelected_Click(object sender, RoutedEventArgs e)
        {
            if (lbVessels2.SelectedItems != null)
            {
                foreach (Vessel vessel in lbVessels2.SelectedItems)
                {
                    vessel.IsSelected = false;
                }
            }
        }

        private void mnuVesselEnableAll_Click(object sender, RoutedEventArgs e)
        {
            if (company.Vessels.Count > 0)
            {
                foreach (Vessel v in company.Vessels)
                {
                    v.IsSelected = true;
                }
            }
        }
        private void mnuVesselDisableAll_Click(object sender, RoutedEventArgs e)
        {
            if (company.Vessels.Count > 0)
            {
                foreach (Vessel v in company.Vessels)
                {
                    v.IsSelected = false;
                }
            }
        }
        #endregion

        #region Validate DGVS  
        //DGV 1
        private void DGV1_CheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (DGV0.SelectedItem != null)
                ((Vessel)DGV0.SelectedItem).NotifyChange("");
        }

        private void DGV1_ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            if (DGV0.SelectedItem != null)
                ((Vessel)DGV0.SelectedItem).NotifyChange("");
        }

        private void DGV1_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Tab)
            {
                if (DGV0.SelectedItem != null)
                    ((Vessel)DGV0.SelectedItem).NotifyChange("");
            }
        }

        private void DGV1_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (DGV0.SelectedItem != null)
                ((Vessel)DGV0.SelectedItem).NotifyChange("");
        }

        private void DGV1_UnloadingRow(object sender, DataGridRowEventArgs e)
        {
            if (DGV0.SelectedItem != null)
                ((Vessel)DGV0.SelectedItem).NotifyChange("");
        }


        //DGV 2
        private void DGV2_ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            Voyage voyage = (Voyage)DGV1.SelectedItem;
            voyage.NotifyChange("");
            if (DGV0.SelectedItem != null)
                ((Vessel)DGV0.SelectedItem).NotifyChange("");
        }

        private void DGV2_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Tab)
            {
                Voyage voyage = (Voyage)DGV1.SelectedItem;
                voyage.NotifyChange("");
                if (DGV0.SelectedItem != null)
                    ((Vessel)DGV0.SelectedItem).NotifyChange("");
            }
        }

        private void DGV2_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Voyage voyage = (Voyage)DGV1.SelectedItem;
            voyage.NotifyChange("");
            if (DGV0.SelectedItem != null)
                ((Vessel)DGV0.SelectedItem).NotifyChange("");
        }

        private void DGV2_UnloadingRow(object sender, DataGridRowEventArgs e)
        {
            if (sender is DataGrid)
            {
                if (DGV1.SelectedItem != null)
                {
                    if (DGV1.SelectedItem is Voyage)
                    {
                        Voyage voyage = (Voyage)DGV1.SelectedItem;
                        voyage.NotifyChange("");
                    }
                }

                if (DGV0.Items.Count > 0 && DGV0.SelectedItem != null)
                    ((Vessel)DGV0.SelectedItem).NotifyChange("");
            }
        }
        #endregion

        #region Graphs
        private void cbVessel_Checked(object sender, RoutedEventArgs e)
        {
            tbPeriod.DataContext = company.Vessels;
            CreateGraphs();
        }
        private void cbVessel_Unchecked(object sender, RoutedEventArgs e)
        {
            tbPeriod.DataContext = company.Vessels;
            CreateGraphs();
        }
        private void cbHasReg_Checked(object sender, RoutedEventArgs e)
        {
            tbPeriod.DataContext = company.Vessels;
            CreateGraphs();
        }
        private void cbHasReg_Unchecked(object sender, RoutedEventArgs e)
        {
            tbPeriod.DataContext = company.Vessels;
            CreateGraphs();
        }

        private void tbPeriod_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            tbPeriod.DataContext = company.Vessels;
            CreateGraphs();
        }
        private void tcMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabControl tc = sender as TabControl;
            if (tc != null)
            {
                TabItem ti = (TabItem)tc.SelectedItem;
                if (ti != null && ti.Header.ToString() == "EEOI Graphs")
                {
                    tbPeriod.DataContext = company.Vessels;
                    CreateGraphs();
                }

                if (ti != null && ti.Header.ToString() == "Voyages")
                {
                    if (DGV0.SelectedIndex == -1)
                    {
                        DGV0.SelectedIndex = 0;
                    }
                }
            }
        }
        private void CreateGraphs()
        {
            GraphEEOI.DataContext = company.Vessels;

            PlotModel myModel = new PlotModel() { Title = "EEOI" };
            //Add curves axis
            var valueXAxis = new LinearAxis() { Position = AxisPosition.Bottom, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Title = "Voyage" };
            var valueYAxis = new LinearAxis() { Position = AxisPosition.Left, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Title = "EEOI [gr CO2/ton x n.mile]" };
            myModel.Axes.Add(valueXAxis);
            myModel.Axes.Add(valueYAxis);

            foreach (Vessel v in company.Vessels)
            {
                if (v.IsSelected)
                {
                    myModel.Series.Add(new LineSeries()
                    {
                        ItemsSource = v.VoyagesEEOIs,
                        Title = v.VesselName
                    });
                }
            }

            foreach (Vessel v in company.Vessels)
            {
                if (v.IsSelected && v.HasReg)
                {
                    int order = 1;
                    //Calculate regression model and add it to curve plot model
                    double[] x = v.VoyagesEEOIs.Select(i => i.X).ToArray();
                    double[] y = v.VoyagesEEOIs.Select(i => i.Y).ToArray();
                    if (x.Length > order)
                    {
                        double[] p = MNet.Fit.Polynomial(x, y, order);
                        myModel.Series.Add(new FunctionSeries(a => MNet.Evaluate.Polynomial(a, p), x[0], x[x.GetUpperBound(0)], 0.1, "Regression Model")
                        {
                            Title = "Trend Line: " + v.VesselName
                        });
                    }
                }
            }
            GraphEEOI.Model = myModel;

            GraphEEOIrv.DataContext = company.Vessels;
            PlotModel myModel1 = new PlotModel() { Title = "EEOI Rolling Average" };
            //Add curves axis
            var valueXAxis1 = new LinearAxis() { Position = AxisPosition.Bottom, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Title = "Voyage" };
            var valueYAxis1 = new LinearAxis() { Position = AxisPosition.Left, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, Title = "EEOI Rolling Average [gr CO2/ton x n.mile]" };
            myModel1.Axes.Add(valueXAxis1);
            myModel1.Axes.Add(valueYAxis1);
            foreach (Vessel v in company.Vessels)
            {
                if (v.IsSelected)
                {
                    myModel1.Series.Add(new LineSeries()
                    {
                        ItemsSource = v.VoyagesEEOIsRA,
                        Title = v.VesselName
                    });
                }
            }

            foreach (Vessel v in company.Vessels)
            {
                if (v.IsSelected && v.HasReg)
                {
                    int order = 1;
                    //Calculate regression model and add it to curve plot model
                    double[] x = v.VoyagesEEOIsRA.Select(i => i.X).ToArray();
                    double[] y = v.VoyagesEEOIsRA.Select(i => i.Y).ToArray();
                    if (x.Length > order)
                    {
                        double[] p = MNet.Fit.Polynomial(x, y, order);
                        myModel1.Series.Add(new FunctionSeries(a => MNet.Evaluate.Polynomial(a, p), x[0], x[x.GetUpperBound(0)], 0.1, "Regression Model")
                        {
                            Title = "Trend Line: " + v.VesselName
                        });
                    }
                }
            }
            GraphEEOIrv.Model = myModel1;
        }
        #endregion

        #region Create Excel
        private void CreateExcelFile()
        {
            //Create Excel Instance.
            Excel.Application xlApp = new Excel.Application();

            //Check if Excel is installed.
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            //Create Excel File using SaveFileDialog.
            string excelFileName = "";
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel|*.xls";

            if (saveFileDialog.ShowDialog() == true)
                excelFileName = saveFileDialog.FileName;
            else
                return;

            //Create Excel WorkBook
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = new Excel.Worksheet();
            object misValue = System.Reflection.Missing.Value;

            //Add a WorkBook.
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            #region CREATE COVER
            //Get the first Worksheet of the active WorkBook.
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Name = "EEOI Report";

            //Header
            xlWorkSheet.Range["B2", "J2"].Merge(false);
            xlWorkSheet.Range["B2", "J2"].Value = company.CompanyName;
            xlWorkSheet.Range["B2", "J2"].RowHeight = 30;
            xlWorkSheet.Range["B2", "J2"].Font.Bold = true;
            xlWorkSheet.Range["B2", "J2"].Font.Size = 22;
            xlWorkSheet.Range["B2", "J2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Range["B2", "J2"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            xlWorkSheet.Range["B2", "J2"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
            xlWorkSheet.Range["B2", "J2"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            xlWorkSheet.Range["B2", "J2"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            xlWorkSheet.Cells[5, 2] = "NAME";
            xlWorkSheet.Cells[5, 3] = "TYPE";
            xlWorkSheet.Cells[5, 4] = "FLAG";
            xlWorkSheet.Cells[5, 5] = "IMO No";
            xlWorkSheet.Cells[5, 6] = "TONNAGE";
            xlWorkSheet.Cells[5, 7] = "Av. EEOI";
            xlWorkSheet.Cells[5, 8] = "EMISSIONS";
            xlWorkSheet.Cells[5, 9] = "TOT. CARGO";
            xlWorkSheet.Cells[5, 10] = "TOT. DIST.";
            xlWorkSheet.Range["B5", "J5"].Font.Bold = true;
            xlWorkSheet.Range["B5", "J5"].Font.Size = 12;
            xlWorkSheet.Range["B5", "J5"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Range["B5", "J5"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            xlWorkSheet.Range["B5", "J5"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
            xlWorkSheet.Range["B5", "J5"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;

            int r = 6;
            foreach (Vessel v in company.Vessels)
            {
                xlWorkSheet.Cells[r, 2] = v.VesselName;
                xlWorkSheet.Cells[r, 3] = v.VesselType.ToString();
                xlWorkSheet.Cells[r, 4] = v.Flag;
                xlWorkSheet.Cells[r, 5] = v.IMO;
                xlWorkSheet.Cells[r, 6] = v.Tonnage;
                xlWorkSheet.Cells[r, 7] = v.VesselEEOI;
                xlWorkSheet.Cells[r, 8] = v.VesselEmissions;
                xlWorkSheet.Cells[r, 9] = v.Voyages.Sum(vo => vo.CargoMass);
                xlWorkSheet.Cells[r, 10] = v.Voyages.Sum(vo => vo.Distance);
                r++;
            }
            r++;
            foreach (Vessel v in company.Vessels)
            {
                xlWorkSheet.Cells[r, 2] = "Vessel Name";
                xlWorkSheet.Cells[r, 3] = "Voyage Id";
                xlWorkSheet.Cells[r, 4] = "Dep. Port";
                xlWorkSheet.Cells[r, 5] = "End Date";
                xlWorkSheet.Cells[r, 6] = "Voy. Type";
                xlWorkSheet.Cells[r, 7] = "EEOI";
                xlWorkSheet.Cells[r, 8] = "Emissions";
                xlWorkSheet.Cells[r, 9] = "Cargo";
                xlWorkSheet.Cells[r, 10] = "Distance";
                xlWorkSheet.Range["B" + r, "J" + r].Font.Bold = true;
                xlWorkSheet.Range["B" + r, "J" + r].Font.Size = 12;
                xlWorkSheet.Range["B" + r, "J" + r].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Range["B" + r, "J" + r].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                xlWorkSheet.Range["B" + r, "J" + r].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
                xlWorkSheet.Range["B" + r, "J" + r].Interior.Pattern = Excel.XlPattern.xlPatternSolid;

                r++;
                xlWorkSheet.Cells[r, 2] = v.VesselName;
                foreach (Voyage vo in v.Voyages)
                {
                    xlWorkSheet.Cells[r, 3] = vo.Id.ToString();
                    xlWorkSheet.Cells[r, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlWorkSheet.Cells[r, 3].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorkSheet.Cells[r, 4] = vo.DeparturePort;
                    xlWorkSheet.Cells[r, 5] = vo.CompletedDate;
                    xlWorkSheet.Cells[r, 6] = vo.VoyageType.ToString();
                    xlWorkSheet.Cells[r, 7] = vo.VoyageEEOI;
                    xlWorkSheet.Cells[r, 8] = vo.VoyageEmissions;
                    xlWorkSheet.Cells[r, 9] = vo.CargoMass;
                    xlWorkSheet.Cells[r, 10] = vo.Distance;
                    r++;
                }
                r++;
            }

            xlWorkSheet.Range["G4", "H" + r].NumberFormat = "0.0";
            xlWorkSheet.Columns.AutoFit();
            #endregion

            #region CREATE REPORT
            //Get the second Worksheet of the active WorkBook.
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            //xlWorkSheet.Name = "Report";
            #endregion

            //Save Excel file.
            xlWorkBook.SaveAs(excelFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            //Release Excel file.
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Excel file created!!");
        }
        #endregion
    }
}