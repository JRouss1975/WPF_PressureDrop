using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml.Serialization;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Windows.Media.Media3D;
using HelixToolkit.Wpf;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MathNet.Numerics.Optimization.LineSearch;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;
using Xceed.Wpf.DataGrid;

namespace WPF_PressureDrop
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<HeadLossCalc> lstHeadLossCalcs { get; set; } = new ObservableCollection<HeadLossCalc>();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = lstHeadLossCalcs;
        }

        private void menuNew_Click(object sender, RoutedEventArgs e)
        {
            lstHeadLossCalcs.Clear();
            txtPCF.Text = "";

            //Draw Model
            Draw();
        }

        private void menuOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            if (_openFileDialog.ShowDialog() == true)
            {
                try
                {
                    XmlSerializer _xmlFormatter = new XmlSerializer(typeof(ObservableCollection<HeadLossCalc>));
                    using (Stream _fileStream = new FileStream(_openFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        _fileStream.Position = 0;
                        lstHeadLossCalcs = (ObservableCollection<HeadLossCalc>)_xmlFormatter.Deserialize(_fileStream);
                    }

                    this.Title = "Piping System Friction Head Loss Calculator v1.08 - " + _openFileDialog.FileName;

                    dgvCalc.ItemsSource = lstHeadLossCalcs;

                    Draw();
                }
                catch (Exception)
                {
                    MessageBox.Show("Please try open the correct file type.");
                }
            }
        }

        private void menuAppend_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            if (_openFileDialog.ShowDialog() == true)
            {
                try
                {
                    XmlSerializer _xmlFormatter = new XmlSerializer(typeof(ObservableCollection<HeadLossCalc>));
                    using (Stream _fileStream = new FileStream(_openFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        _fileStream.Position = 0;
                        ObservableCollection<HeadLossCalc> lstAppendedHeadLossCalcs = new ObservableCollection<HeadLossCalc>();
                        lstAppendedHeadLossCalcs = (ObservableCollection<HeadLossCalc>)_xmlFormatter.Deserialize(_fileStream);

                        foreach (HeadLossCalc item in lstAppendedHeadLossCalcs)
                        {
                            lstHeadLossCalcs.Add(item);
                        }
                    }

                    this.Title = "Piping System Friction Head Loss Calculator v1.08 - " + _openFileDialog.FileName;

                    dgvCalc.ItemsSource = lstHeadLossCalcs;

                    Draw();
                }
                catch (Exception)
                {
                    MessageBox.Show("Please try open the correct file type.");
                }
            }
        }

        private void menuSave_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog _saveFileDialog = new SaveFileDialog();
            if (_saveFileDialog.ShowDialog() == true)
            {
                XmlSerializer _xmlFormatter = new XmlSerializer(typeof(ObservableCollection<HeadLossCalc>));
                using (Stream _fileStream = new FileStream(_saveFileDialog.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    _fileStream.Position = 0;
                    _xmlFormatter.Serialize(_fileStream, lstHeadLossCalcs);
                }
            }
        }

        private void menuImport_Click(object sender, RoutedEventArgs e)
        {
            //Local Variables
            List<string> lstLines = new List<string>();
            List<string> selectedLines = new List<string>();
            string[] Headers = new string[] { "WELD", "PIPE", "FLANGE", "FLANGE-BLIND", "GASKET", "END-CONNECTION-PIPELINE", "VALVE", "INDUCTION-START", "BEND", "INDUCTION-END", "MESSAGE-ROUND", "REDUCER-CONCENTRIC" };

            Double.TryParse(txtTemp.Text, out double t);
            Double.TryParse(txte.Text, out double ep);
            Double.TryParse(txtQ.Text, out double q);
            double temperature = t;
            double epsilon = ep;
            double qh = q;

            //Open File
            OpenFileDialog _openFileDialog = new OpenFileDialog();
            if (_openFileDialog.ShowDialog() == true)
            {
                using (Stream _fileStream = new FileStream(_openFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    string txt = "";
                    using (var streamReader = new StreamReader(_fileStream, Encoding.UTF8, true, 128))
                    {
                        String line;
                        int n = 0;
                        while ((line = streamReader.ReadLine()) != null)
                        {
                            lstLines.Add(line);
                            n = ++n;
                            txt = txt + (n + "... " + line + "\n");
                        }
                        txtPCF.Text = txt;
                    }
                }
            }

            //Search for Items
            for (int i = 0; i < lstLines.Count - 1; i++)
            {

                if (lstLines[i] == "PIPE")
                {
                    selectedLines.Add(lstLines[i]);

                    HeadLossCalc pd = new HeadLossCalc(i + 1);
                    pd.ElementType = ItemType.Pipe;
                    pd.t = temperature;
                    pd.epsilon = epsilon;
                    pd.qh = qh;

                    int c = 1;
                    bool endConditionFound = false;

                    string str = "";
                    List<string> p1 = new List<string>();
                    List<string> p2 = new List<string>();
                    List<string> w0 = new List<string>();

                    do
                    {
                        if ((i + c) > lstLines.Count - 1) break;
                        str = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList().First();

                        if (str == "END-POINT")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            p1 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            selectedLines.Add(lstLines[i + c + 1]);
                            p2 = lstLines[i + c + 1].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            i = i + 2;
                            pd.Node1.X = Double.Parse(p1[1]);
                            pd.Node1.Y = Double.Parse(p1[2]);
                            pd.Node1.Z = Double.Parse(p1[3]);
                            pd.Node2.X = Double.Parse(p2[1]);
                            pd.Node2.Y = Double.Parse(p2[2]);
                            pd.Node2.Z = Double.Parse(p2[3]);
                            pd.d = Double.Parse(p1[4]);
                            pd.L = pd.Distance(pd.Node1.X, pd.Node2.X, pd.Node1.Y, pd.Node2.Y, pd.Node1.Z, pd.Node2.Z) / 1000;
                        }

                        if (str == "WEIGHT")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.W = Double.Parse(w0[1]) * pd.L;
                        }

                        if (Headers.Contains(str))
                        {
                            endConditionFound = true;
                        }

                        c++;
                    } while (!endConditionFound);

                    lstHeadLossCalcs.Add(pd);

                }

                if (lstLines[i] == "REDUCER-CONCENTRIC")
                {
                    selectedLines.Add(lstLines[i]);

                    HeadLossCalc pd = new HeadLossCalc(i + 1);
                    pd.t = temperature;
                    pd.epsilon = epsilon;
                    pd.qh = qh;

                    int c = 1;
                    bool endConditionFound = false;

                    string str = "";
                    double d1 = 0;
                    double d2 = 0;
                    List<string> p1 = new List<string>();
                    List<string> p2 = new List<string>();
                    List<string> w0 = new List<string>();

                    do
                    {
                        if ((i + c) > lstLines.Count - 1) break;
                        str = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList().First();

                        if (str == "END-POINT")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            p1 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            selectedLines.Add(lstLines[i + c + 1]);
                            p2 = lstLines[i + c + 1].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            i = i + 2;
                            pd.Node1.X = Double.Parse(p1[1]);
                            pd.Node1.Y = Double.Parse(p1[2]);
                            pd.Node1.Z = Double.Parse(p1[3]);
                            pd.Node2.X = Double.Parse(p2[1]);
                            pd.Node2.Y = Double.Parse(p2[2]);
                            pd.Node2.Z = Double.Parse(p2[3]);
                            pd.d = Double.Parse(p1[4]);
                            d1 = Double.Parse(p1[4]);
                            d2 = Double.Parse(p2[4]);
                            pd.L = pd.Distance(pd.Node1.X, pd.Node2.X, pd.Node1.Y, pd.Node2.Y, pd.Node1.Z, pd.Node2.Z) / 1000;

                            if (d1 < d2)
                            {
                                pd.ElementType = ItemType.Expander;
                                pd.d = Double.Parse(p1[4]);
                                pd.d1 = d1;
                                pd.d2 = d2;
                            }
                            else
                            {
                                pd.ElementType = ItemType.Reducer;
                                pd.d = Double.Parse(p1[4]);
                                pd.d1 = d2;
                                pd.d2 = d1;
                            }
                        }

                        if (str == "WEIGHT")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.W = Double.Parse(w0[1]);
                        }


                        if (Headers.Contains(str))
                        {
                            endConditionFound = true;
                        }

                        c++;
                    } while (!endConditionFound);

                    lstHeadLossCalcs.Add(pd);
                }

                if (lstLines[i] == "BEND")
                {
                    selectedLines.Add(lstLines[i]);

                    HeadLossCalc pd = new HeadLossCalc(i + 1);
                    pd.ElementType = ItemType.Bend;
                    pd.t = temperature;
                    pd.epsilon = epsilon;
                    pd.qh = qh;

                    int c = 1;
                    bool endConditionFound = false;

                    string str = "";
                    List<string> p1 = new List<string>();
                    List<string> p2 = new List<string>();
                    List<string> c1 = new List<string>();
                    List<string> a0 = new List<string>();
                    List<string> r0 = new List<string>();
                    List<string> w0 = new List<string>();
                    double a = 0;
                    double r = 0;

                    do
                    {
                        if ((i + c) > lstLines.Count - 1) break;
                        str = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList().First();

                        if (str == "END-POINT")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            p1 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            selectedLines.Add(lstLines[i + c + 1]);
                            p2 = lstLines[i + c + 1].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            i = i + 2;
                            pd.Node1.X = Double.Parse(p1[1]);
                            pd.Node1.Y = Double.Parse(p1[2]);
                            pd.Node1.Z = Double.Parse(p1[3]);
                            pd.Node2.X = Double.Parse(p2[1]);
                            pd.Node2.Y = Double.Parse(p2[2]);
                            pd.Node2.Z = Double.Parse(p2[3]);
                            pd.d = Double.Parse(p1[4]);
                        }

                        if (str == "ANGLE")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            a0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            a = Double.Parse(a0[1]);
                            pd.a = a / 100;
                            pd.n = (a / 100) / 90;
                        }

                        if (str == "BEND-RADIUS")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            r0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            r = Double.Parse(r0[1]);
                            pd.r = r;
                            pd.L = (r / 1000) * (Math.PI * (a / 100) / 180);
                        }

                        if (str == "WEIGHT")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.W = Double.Parse(w0[1]);
                        }

                        if (Headers.Contains(str))
                        {
                            endConditionFound = true;
                        }

                        c++;
                    } while (!endConditionFound);

                    lstHeadLossCalcs.Add(pd);
                }

                if (lstLines[i] == "VALVE")
                {
                    selectedLines.Add(lstLines[i]);

                    HeadLossCalc pd = new HeadLossCalc(i + 1);
                    pd.ElementType = ItemType.Butterfly;
                    pd.t = temperature;
                    pd.epsilon = epsilon;
                    pd.qh = qh;

                    int c = 1;
                    bool endConditionFound = false;

                    string str = "";
                    List<string> p1 = new List<string>();
                    List<string> p2 = new List<string>();
                    List<string> w0 = new List<string>();

                    do
                    {
                        if ((i + c) > lstLines.Count - 1) break;
                        str = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList().First();

                        if (str == "END-POINT")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            p1 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            selectedLines.Add(lstLines[i + c + 1]);
                            p2 = lstLines[i + c + 1].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            i = i + 2;
                            pd.Node1.X = Double.Parse(p1[1]);
                            pd.Node1.Y = Double.Parse(p1[2]);
                            pd.Node1.Z = Double.Parse(p1[3]);
                            pd.Node2.X = Double.Parse(p2[1]);
                            pd.Node2.Y = Double.Parse(p2[2]);
                            pd.Node2.Z = Double.Parse(p2[3]);
                            pd.d = Double.Parse(p1[4]);
                            pd.L = pd.Distance(pd.Node1.X, pd.Node2.X, pd.Node1.Y, pd.Node2.Y, pd.Node1.Z, pd.Node2.Z) / 1000;
                        }

                        if (str == "WEIGHT")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.W = Double.Parse(w0[1]);
                        }

                        if (Headers.Contains(str))
                        {
                            endConditionFound = true;
                        }

                        c++;
                    } while (!endConditionFound);

                    lstHeadLossCalcs.Add(pd);
                }

                if (lstLines[i] == "FLANGE")
                {
                    selectedLines.Add(lstLines[i]);

                    HeadLossCalc pd = new HeadLossCalc(i + 1);
                    pd.ElementType = ItemType.Flange;
                    pd.t = temperature;
                    pd.epsilon = epsilon;
                    pd.qh = qh;

                    int c = 1;
                    bool endConditionFound = false;

                    string str = "";
                    List<string> p1 = new List<string>();
                    List<string> p2 = new List<string>();
                    List<string> w0 = new List<string>();

                    do
                    {
                        if ((i + c) > lstLines.Count - 1) break;
                        str = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList().First();

                        if (str == "END-POINT")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            p1 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            selectedLines.Add(lstLines[i + c + 1]);
                            p2 = lstLines[i + c + 1].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            i = i + 2;
                            pd.Node1.X = Double.Parse(p1[1]);
                            pd.Node1.Y = Double.Parse(p1[2]);
                            pd.Node1.Z = Double.Parse(p1[3]);
                            pd.Node2.X = Double.Parse(p2[1]);
                            pd.Node2.Y = Double.Parse(p2[2]);
                            pd.Node2.Z = Double.Parse(p2[3]);
                            pd.d = Double.Parse(p1[4]);
                            pd.L = pd.Distance(pd.Node1.X, pd.Node2.X, pd.Node1.Y, pd.Node2.Y, pd.Node1.Z, pd.Node2.Z) / 1000;
                        }

                        if (str == "WEIGHT")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.W = Double.Parse(w0[1]);
                        }

                        if (str == "ITEM-DESCRIPTION")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.Notes = w0[2] + w0[3];
                        }

                        if (Headers.Contains(str))
                        {
                            endConditionFound = true;
                        }

                        c++;
                    } while (!endConditionFound);

                    lstHeadLossCalcs.Add(pd);

                }

                if (lstLines[i] == "FLANGE-BLIND")
                {
                    selectedLines.Add(lstLines[i]);

                    HeadLossCalc pd = new HeadLossCalc(i + 1);
                    pd.ElementType = ItemType.FlangeBlind;
                    pd.t = temperature;
                    pd.epsilon = epsilon;
                    pd.qh = qh;

                    int c = 1;
                    bool endConditionFound = false;

                    string str = "";
                    List<string> p1 = new List<string>();
                    List<string> p2 = new List<string>();
                    List<string> w0 = new List<string>();

                    do
                    {
                        if ((i + c) > lstLines.Count - 1) break;
                        str = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList().First();

                        if (str == "END-POINT")
                        {
                            selectedLines.Add(lstLines[i + c]);
                            p1 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            selectedLines.Add(lstLines[i + c + 1]);
                            p2 = lstLines[i + c + 1].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            i = i + 2;
                            pd.Node1.X = Double.Parse(p1[1]);
                            pd.Node1.Y = Double.Parse(p1[2]);
                            pd.Node1.Z = Double.Parse(p1[3]);
                            pd.Node2.X = Double.Parse(p2[1]);
                            pd.Node2.Y = Double.Parse(p2[2]);
                            pd.Node2.Z = Double.Parse(p2[3]);
                            pd.d = Double.Parse(p1[4]);
                            pd.L = pd.Distance(pd.Node1.X, pd.Node2.X, pd.Node1.Y, pd.Node2.Y, pd.Node1.Z, pd.Node2.Z) / 1000;
                        }

                        if (str == "WEIGHT")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.W = Double.Parse(w0[1]);
                        }

                        if (str == "ITEM-DESCRIPTION")
                        {
                            w0 = lstLines[i + c].Split(' ').ToList<string>().Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                            pd.Notes = w0[2] + w0[3];
                        }

                        if (Headers.Contains(str))
                        {
                            endConditionFound = true;
                        }

                        c++;
                    } while (!endConditionFound);

                    lstHeadLossCalcs.Add(pd);

                }

            }

            //Assing Line Tag
            string lineTag = lstLines[6].Split().Last();
            foreach (HeadLossCalc item in lstHeadLossCalcs)
            {
                if (item.Line == "")
                    item.Line = lineTag;
            }

            //Draw Model
            Draw();
        }

        private void menuExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void menuRemoveItem_Click(object sender, RoutedEventArgs e)
        {
            int c = 0;
            MenuItem mnu = sender as MenuItem;
            DataGridControl dgv = null;
            if (mnu != null)
            {
                dgv = ((ContextMenu)mnu.Parent).PlacementTarget as DataGridControl;
            }

            if (dgv.Items.Count > 0 && dgv.SelectedIndex > -1)
            {
                var selection = dgv.SelectedItems;

                List<HeadLossCalc> lstSelection = new List<HeadLossCalc>();
                foreach (var item in selection)
                {
                    lstSelection.Add((HeadLossCalc)item);
                }

                foreach (HeadLossCalc item in lstSelection)
                {
                    lstHeadLossCalcs.Remove(item);
                    c++;
                }
                MessageBox.Show($"{c} items removed!");
            }
            else
            {
                MessageBox.Show("No item to remove!");
            }
        }

        private void menuAddItem_Click(object sender, RoutedEventArgs e)
        {
            Double.TryParse(txtTemp.Text, out double t);
            Double.TryParse(txte.Text, out double ep);
            Double.TryParse(txtQ.Text, out double q);

            int maxId = 0;

            if (lstHeadLossCalcs.Count > 0)
                maxId = lstHeadLossCalcs.Max(x => x.ElementId);


            if (dgvCalc.SelectedIndex == lstHeadLossCalcs.Count - 1 || dgvCalc.SelectedIndex == -1)
            {
                lstHeadLossCalcs.Add(new HeadLossCalc() { ElementId = maxId + 1, t = t, epsilon = ep, qh = q });
                return;
            }


            if (dgvCalc.SelectedIndex > -1)
            {
                lstHeadLossCalcs.Insert(dgvCalc.SelectedIndex, new HeadLossCalc() { ElementId = maxId + 1, t = t, epsilon = ep, qh = q });
            }


        }

        private void dgvCalc_SelectionChanged(object sender, Xceed.Wpf.DataGrid.DataGridSelectionChangedEventArgs e)
        {
            HeadLossCalc selectedItem = (HeadLossCalc)e.SelectionInfos[0].DataGridContext.CurrentItem;
            if (selectedItem != null)
            {
                double totalHl = 0;
                double totalV = 0;
                double totalL = 0;
                var selection = dgvCalc.SelectedItems;
                foreach (HeadLossCalc hlc in selection)
                {
                    totalL = totalL + hlc.L;
                    totalV = totalV + hlc.Vm;
                    totalHl = totalHl + hlc.hLm;

                }
                tbL.Text = totalL.ToString("N3");
                tbV.Text = totalV.ToString("N3");
                tbHlmm.Text = totalHl.ToString("N2");
            }
        }

        private void btnApplyAll_Click(object sender, RoutedEventArgs e)
        {
            Double.TryParse(txtTemp.Text, out double t);
            Double.TryParse(txte.Text, out double ep);
            Double.TryParse(txtQ.Text, out double q);
            Double.TryParse(txtL.Text, out double l);
            Double.TryParse(txtD.Text, out double d);
            foreach (HeadLossCalc hlc in lstHeadLossCalcs)
            {
                if (cbTemp.IsChecked == true) hlc.t = t;
                if (cbe.IsChecked == true) hlc.epsilon = ep;
                if (cbQ.IsChecked == true) hlc.qh = q;
                if (cbL.IsChecked == true) hlc.L = l;
                if (cbD.IsChecked == true) hlc.d = d;
            }
        }

        private void btnApplySelection_Click(object sender, RoutedEventArgs e)
        {
            Double.TryParse(txtTemp.Text, out double t);
            Double.TryParse(txte.Text, out double ep);
            Double.TryParse(txtQ.Text, out double q);
            Double.TryParse(txtL.Text, out double l);
            Double.TryParse(txtD.Text, out double d);
            var selection = dgvCalc.SelectedItems;
            foreach (HeadLossCalc hlc in selection)
            {
                if (cbTemp.IsChecked == true) hlc.t = t;
                if (cbe.IsChecked == true) hlc.epsilon = ep;
                if (cbQ.IsChecked == true) hlc.qh = q;
                if (cbL.IsChecked == true) hlc.L = l;
                if (cbD.IsChecked == true) hlc.d = d;
            }
        }

        private void menuSortItem_Click(object sender, RoutedEventArgs e)
        {
            if (lstHeadLossCalcs.Count > 0)
            {
                var items = lstHeadLossCalcs.OrderBy(x => x.ElementId).ToList();
                lstHeadLossCalcs.Clear();

                foreach (var item in items)
                {
                    lstHeadLossCalcs.Add(item);
                }
            }
        }

        private void menuRenumberItem_Click(object sender, RoutedEventArgs e)
        {
            int i = -1;

            MenuItem mnu = sender as MenuItem;
            DataGridControl dgv = null;
            if (mnu != null)
            {
                dgv = ((ContextMenu)mnu.Parent).PlacementTarget as DataGridControl;
            }

            if (dgv.Name.ToString() == "dgvCalc")
            {
                Int32.TryParse(tbNumber.Text, out i);
            }
            if (dgv.Name.ToString() == "dgvBOM")
            {
                Int32.TryParse(tbNumber1.Text, out i);
            }

            var selection = dgv.SelectedItems;
            if (selection.Count > 0)
            {
                foreach (HeadLossCalc hlc in selection)
                {
                    hlc.ElementId = i++;
                }
            }

            //Sort list
            var temp = lstHeadLossCalcs.OrderBy(x => x.ElementId).ToList();
            lstHeadLossCalcs.Clear();
            foreach (HeadLossCalc h in temp)
            {
                lstHeadLossCalcs.Add(h);
            }

        }

        private List<HeadLossCalc> copiedItems;
        private void menuCopyItem_Click(object sender, RoutedEventArgs e)
        {
            copiedItems = new List<HeadLossCalc>();
            var selection = dgvCalc.SelectedItems;
            if (selection.Count > 0)
            {
                foreach (HeadLossCalc hlc in selection)
                {
                    copiedItems.Add(hlc);
                }
            }
        }

        private void menuPasteItem_Click(object sender, RoutedEventArgs e)
        {
            int pos = dgvCalc.SelectedIndex;
            if (copiedItems.Count > 0 && pos > -1)
            {
                foreach (HeadLossCalc hlc in copiedItems)
                {
                    HeadLossCalc newItem = (HeadLossCalc)hlc.Clone();

                    newItem.Node1 = new Node();
                    newItem.Node1.X = hlc.Node1.X;
                    newItem.Node1.Y = hlc.Node1.Y;
                    newItem.Node1.Z = hlc.Node1.Z;

                    newItem.Node2 = new Node();
                    newItem.Node2.X = hlc.Node2.X;
                    newItem.Node2.Y = hlc.Node2.Y;
                    newItem.Node2.Z = hlc.Node2.Z;

                    lstHeadLossCalcs.Insert(++pos, newItem);
                }
            }
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            Draw();
        }

        public void Draw()
        {
            //3D Drawing
            BillboardTextVisual3D txt1;
            LinesVisual3D linesVisual;
            PointsVisual3D pointsVisual;
            Point3DCollection pts = new Point3DCollection();

            View1.Children.Clear();

            foreach (HeadLossCalc i in lstHeadLossCalcs)
            {
                Point3D p1 = new Point3D(i.Node1.X, i.Node1.Y, i.Node1.Z);
                pts.Add(p1);
                Point3D p2 = new Point3D(i.Node2.X, i.Node2.Y, i.Node2.Z);
                pts.Add(p2);
                txt1 = new BillboardTextVisual3D();
                txt1.Text = i.ElementId.ToString();
                txt1.Position = new Point3D((i.Node1.X + i.Node2.X) / 2, (i.Node1.Y + i.Node2.Y) / 2, (i.Node1.Z + i.Node2.Z) / 2);
                View1.Children.Add(txt1);
            }

            GridLinesVisual3D grid = new GridLinesVisual3D();
            grid.Length = 50000;
            grid.Width = 50000;
            grid.MajorDistance = 10000;
            grid.MinorDistance = 1000;
            grid.Visible = true;
            grid.Thickness = 10;
            View1.Children.Add(grid);

            pointsVisual = new PointsVisual3D { Color = Colors.Red, Size = 4 };
            pointsVisual.Points = pts;
            View1.Children.Add(pointsVisual);
            linesVisual = new LinesVisual3D { Color = Colors.Blue };
            linesVisual.Points = pts;
            linesVisual.Thickness = 2;
            View1.Children.Add(linesVisual);
            View1.ZoomExtents(10);

        }

        private void View1_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var viewport = (HelixViewport3D)sender;
            var firstHit = viewport.Viewport.FindHits(e.GetPosition(viewport)).FirstOrDefault();
            if (firstHit != null)
            {
                if (firstHit.Visual is BillboardTextVisual3D)
                {
                    BillboardTextVisual3D t = (BillboardTextVisual3D)firstHit.Visual;
                    dgvCalc.CurrentItem = lstHeadLossCalcs.FirstOrDefault(x => x.ElementId == Convert.ToInt32(t.Text.ToString()));
                }
            }
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void menuExportExcel_Click(object sender, RoutedEventArgs e)
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
            Excel.Worksheet xlWorkSheet;// = new Excel.Worksheet();
            object misValue = System.Reflection.Missing.Value;

            //Add a WorkBook.
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            #region CREATE COVER

            ////Get the first Worksheet of the active WorkBook.
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //xlWorkSheet.Name = "Cover";

            ////Header
            //xlWorkSheet.Range["B3", "J3"].Merge(false);
            //xlWorkSheet.Range["B3", "J3"].Value = DataProvider.ProjectInfo.ProjectCategory;
            //xlWorkSheet.Range["B3", "J3"].RowHeight = 35;
            //xlWorkSheet.Range["B3", "J3"].Font.Bold = true;
            //xlWorkSheet.Range["B3", "J3"].Font.Size = 22;
            //xlWorkSheet.Range["B3", "J3"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Range["B3", "J3"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //xlWorkSheet.Range["B3", "J3"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //xlWorkSheet.Range["B3", "J3"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["B3", "J3"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            //xlWorkSheet.Range["B4", "B20"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //xlWorkSheet.Range["B4", "B20"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["B4", "B20"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            //xlWorkSheet.Range["J4", "J20"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //xlWorkSheet.Range["J4", "J20"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["J4", "J20"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            //xlWorkSheet.Range["B21", "J21"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //xlWorkSheet.Range["B21", "J21"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["B21", "J21"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //xlWorkSheet.Range["B21", "J21"].RowHeight = 35;

            ////Name
            //xlWorkSheet.Cells[10, 3].Value = "ΟΝΟΜΑ:";
            //xlWorkSheet.Cells[10, 3].RowHeight = 14;
            //xlWorkSheet.Cells[10, 3].Font.Bold = true;

            //xlWorkSheet.Range["C11", "I11"].Merge(false);
            //xlWorkSheet.Range["C11", "I11"].Value = DataProvider.ProjectInfo.CustomerName;
            //xlWorkSheet.Range["C11", "I11"].RowHeight = 24;
            //xlWorkSheet.Range["C11", "I11"].Font.Bold = true;
            //xlWorkSheet.Range["C11", "I11"].Font.Size = 18;
            //xlWorkSheet.Range["C11", "I11"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Range["C11", "I11"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //xlWorkSheet.Range["C11", "I11"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Range["C11", "I11"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["C11", "I11"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            ////Address
            //xlWorkSheet.Cells[14, 3].Value = "ΔΙΕΥΘΥΝΣΗ ΕΡΓΟΥ:";
            //xlWorkSheet.Cells[14, 3].RowHeight = 14;
            //xlWorkSheet.Cells[14, 3].Font.Bold = true;

            //xlWorkSheet.Range["C15", "I15"].Merge(false);
            //xlWorkSheet.Range["C15", "I15"].Value = DataProvider.ProjectInfo.ProjectAddress + " - " + DataProvider.ProjectInfo.ProjectCity;
            //xlWorkSheet.Range["C15", "I15"].RowHeight = 22.5;
            //xlWorkSheet.Range["C15", "I15"].Font.Bold = true;
            //xlWorkSheet.Range["C15", "I15"].Font.Size = 14;
            //xlWorkSheet.Range["C15", "I15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Range["C15", "I15"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //xlWorkSheet.Range["C15", "I15"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Range["C15", "I15"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["C15", "I15"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            ////Date
            //xlWorkSheet.Cells[17, 3].Value = "ΗΜΕΡΟΜΗΝΙΑ:";
            //xlWorkSheet.Cells[17, 3].RowHeight = 14;
            //xlWorkSheet.Cells[17, 3].Font.Bold = true;

            //xlWorkSheet.Range["C18", "I18"].Merge(false);
            //xlWorkSheet.Range["C18", "I18"].Value = DataProvider.ProjectInfo.Date.ToString("MMMM") + " " + DataProvider.ProjectInfo.Date.Year;
            //xlWorkSheet.Range["C18", "I18"].RowHeight = 22.5;
            //xlWorkSheet.Range["C18", "I18"].Font.Bold = true;
            //xlWorkSheet.Range["C18", "I18"].Font.Size = 14;
            //xlWorkSheet.Range["C18", "I18"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Range["C18", "I18"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //xlWorkSheet.Range["C18", "I18"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Range["C18", "I18"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["C18", "I18"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            //#endregion CREATE COVER

            //#region CREATE REPORT

            ////Get the second Worksheet of the active WorkBook.
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            //xlWorkSheet.Name = "Report";

            ////Excel.Worksheet xlNewSheet = null;
            ////xlNewSheet = xlWorkBook.Worksheets.Add(xlWorkSheet, Type.Missing, misValue, misValue);
            ////xlNewSheet.Name = "New";

            ////Set all cells font size.
            //xlWorkSheet.Cells.Font.Size = 11;

            //int r = 0;
            //List<int> titlePos = new List<int>();
            //List<int> headerPos = new List<int>();

            //for (int i = 0; i < DataProvider.Measurements.Count; i++)
            //{
            //    Measurement m = DataProvider.Measurements[i];

            //    r = r + 1;

            //    if (m.Header != "")
            //    {
            //        r += 3;
            //        xlWorkSheet.Cells[r, 2] = m.Header;
            //        xlWorkSheet.Cells[r, 2].Font.Bold = true;
            //        xlWorkSheet.Cells[r, 2].Font.Size = 20;
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].Merge(false);
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //        headerPos.Add(r);
            //    }

            //    if (m.Title != "")
            //    {
            //        r = r + 2;
            //        xlWorkSheet.Cells[r, 2] = m.Title;
            //        xlWorkSheet.Cells[r, 2].Font.Bold = true;
            //        xlWorkSheet.Cells[r, 2].Font.Size = 16;
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].Merge(false);
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //        xlWorkSheet.Range[$"b{r}", $"k{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //        titlePos.Add(r);
            //    }

            //    if (m.Category != "")
            //    {
            //        xlWorkSheet.Cells[r, 2] = m.Category;
            //    }

            //    switch (m.MeasureType)
            //    {
            //        case MeasureType.M:
            //            if (m.Q == 1)
            //            {
            //                if (m.IsNegative)
            //                {
            //                    xlWorkSheet.Cells[r, 11] = -m.D1;
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                    xlWorkSheet.Range[$"a{r}", $"k{r}"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //                }
            //                else
            //                {
            //                    xlWorkSheet.Cells[r, 11] = m.D1;
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                }
            //            }
            //            else
            //            {
            //                xlWorkSheet.Cells[r, 8] = m.D1;
            //                xlWorkSheet.Cells[r, 8].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 9] = "x";
            //                xlWorkSheet.Cells[r, 10] = m.Q;
            //                if (m.IsNegative)
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=-H{r}*J{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                    xlWorkSheet.Range[$"a{r}", $"k{r}"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //                }
            //                else
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=H{r}*J{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                }
            //            }

            //            xlWorkSheet.Cells[r, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Range[$"c{r}", $"j{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Cells[r, 11].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            break;

            //        case MeasureType.M2:
            //            if (m.Q == 1)
            //            {
            //                xlWorkSheet.Cells[r, 4] = m.D1;
            //                xlWorkSheet.Cells[r, 4].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 5] = "x";
            //                xlWorkSheet.Cells[r, 6] = m.D2;
            //                xlWorkSheet.Cells[r, 6].NumberFormat = "###0.00";
            //                if (m.IsNegative)
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=-D{r}*F{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                    xlWorkSheet.Range[$"a{r}", $"k{r}"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //                }
            //                else
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=D{r}*F{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                }
            //            }
            //            else
            //            {
            //                xlWorkSheet.Cells[r, 4] = m.D1;
            //                xlWorkSheet.Cells[r, 4].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 5] = "x";
            //                xlWorkSheet.Cells[r, 6] = m.D2;
            //                xlWorkSheet.Cells[r, 6].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 7] = "x";
            //                xlWorkSheet.Cells[r, 8] = m.Q;
            //                if (m.IsNegative)
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=-D{r}*F{r}*H{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                    xlWorkSheet.Range[$"a{r}", $"k{r}"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //                }
            //                else
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=D{r}*F{r}*H{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                }

            //                xlWorkSheet.Cells[r, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //                xlWorkSheet.Range[$"c{r}", $"j{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //                xlWorkSheet.Cells[r, 11].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            }

            //            xlWorkSheet.Cells[r, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Range[$"c{r}", $"j{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Cells[r, 11].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            break;

            //        case MeasureType.M3:
            //            if (m.Q == 1)
            //            {
            //                xlWorkSheet.Cells[r, 4] = m.D1;
            //                xlWorkSheet.Cells[r, 4].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 5] = "x";
            //                xlWorkSheet.Cells[r, 6] = m.D2;
            //                xlWorkSheet.Cells[r, 6].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 7] = "x";
            //                xlWorkSheet.Cells[r, 8] = m.D3;
            //                xlWorkSheet.Cells[r, 8].NumberFormat = "###0.00";
            //                if (m.IsNegative)
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=-D{r}*F{r}*H{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                    xlWorkSheet.Range[$"a{r}", $"k{r}"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //                }
            //                else
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=D{r}*F{r}*H{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                }
            //            }
            //            else
            //            {
            //                xlWorkSheet.Cells[r, 4] = m.D1;
            //                xlWorkSheet.Cells[r, 4].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 5] = "x";
            //                xlWorkSheet.Cells[r, 6] = m.D2;
            //                xlWorkSheet.Cells[r, 6].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 7] = "x";
            //                xlWorkSheet.Cells[r, 8] = m.D3;
            //                xlWorkSheet.Cells[r, 8].NumberFormat = "###0.00";
            //                xlWorkSheet.Cells[r, 9] = "x";
            //                xlWorkSheet.Cells[r, 10] = m.Q;
            //                if (m.IsNegative)
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=-D{r}*F{r}*H{r}*J{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                    xlWorkSheet.Range[$"a{r}", $"k{r}"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //                }
            //                else
            //                {
            //                    xlWorkSheet.Cells[r, 11].Formula = $"=D{r}*F{r}*H{r}*J{r}";
            //                    xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                }
            //            }

            //            xlWorkSheet.Cells[r, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Range[$"c{r}", $"j{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Cells[r, 11].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            break;

            //        case MeasureType.Cross:
            //            xlWorkSheet.Cells[r, 4] = m.D1;
            //            xlWorkSheet.Cells[r, 4].NumberFormat = "###0.00";
            //            xlWorkSheet.Cells[r, 5] = "+";
            //            xlWorkSheet.Cells[r, 6] = m.D2;
            //            xlWorkSheet.Cells[r, 6].NumberFormat = "###0.00";
            //            xlWorkSheet.Cells[r, 7] = "x";
            //            xlWorkSheet.Cells[r, 8] = m.D3;
            //            xlWorkSheet.Cells[r, 8].NumberFormat = "###0.00";
            //            xlWorkSheet.Cells[r, 9] = "x";
            //            xlWorkSheet.Cells[r, 10] = m.Q;

            //            if (m.IsNegative)
            //            {
            //                xlWorkSheet.Cells[r, 11].Formula = $"=-(D{r}+F{r})*H{r}*J{r}";
            //                xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //                xlWorkSheet.Range[$"a{r}", $"k{r}"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //            }
            //            else
            //            {
            //                xlWorkSheet.Cells[r, 11].Formula = $"=(D{r}+F{r})*H{r}*J{r}";
            //                xlWorkSheet.Cells[r, 11].NumberFormat = "###0.00";
            //            }

            //            xlWorkSheet.Cells[r, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Range[$"c{r}", $"j{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            xlWorkSheet.Cells[r, 11].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //            break;

            //        default:
            //            break;
            //    }

            //    if (m.SubTitle != "")
            //    {
            //        xlWorkSheet.Cells[r, 2] = m.SubTitle;
            //        xlWorkSheet.Cells[r, 2].Font.Bold = true;
            //        xlWorkSheet.Cells[r, 2].Font.Size = 12;
            //        xlWorkSheet.Cells[r, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //        //xlWorkSheet.Range[$"b{r}", $"k{r}"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //    }
            //}
            //titlePos.Add(r + 3);

            ////Set Columns Width.
            //xlWorkSheet.Columns[1].ColumnWidth = 10;
            //xlWorkSheet.Columns[2].ColumnWidth = 20;
            //xlWorkSheet.Columns[3].ColumnWidth = 1.6;
            //xlWorkSheet.Columns[4].ColumnWidth = 7;
            //xlWorkSheet.Columns[5].ColumnWidth = 1.6;
            //xlWorkSheet.Columns[6].ColumnWidth = 7;
            //xlWorkSheet.Columns[7].ColumnWidth = 1.6;
            //xlWorkSheet.Columns[8].ColumnWidth = 7;
            //xlWorkSheet.Columns[9].ColumnWidth = 1.6;
            //xlWorkSheet.Columns[10].ColumnWidth = 3;
            //xlWorkSheet.Columns[11].ColumnWidth = 8;

            ////Set Columns Alignment.
            //xlWorkSheet.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            //xlWorkSheet.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[9].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Columns[11].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            ////Create Sums
            //for (int i = 1; i < titlePos.Count; i++)
            //{
            //    int sp = titlePos[i - 1] + 1;
            //    int ep = titlePos[i] - 2;

            //    if (headerPos.Contains(ep - 1))
            //        ep = ep - 4;

            //    xlWorkSheet.Cells[ep, 9] = "ΣΥΝΟΛΟ:";
            //    xlWorkSheet.Cells[ep, 9].Font.Bold = true;
            //    xlWorkSheet.Cells[ep, 9].Font.Size = 12;
            //    xlWorkSheet.Cells[ep, 11].Formula = $"=SUM(k{sp}:k{ep - 1})";

            //    xlWorkSheet.Cells[ep, 11].NumberFormat = "###0.00";
            //    xlWorkSheet.Cells[ep, 11].Font.Bold = true;
            //    xlWorkSheet.Cells[ep, 11].Font.Size = 12;
            //    xlWorkSheet.Cells[ep, 11].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            //}

            ////Delete first two rows
            //xlWorkSheet.Rows[1, misValue].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            //xlWorkSheet.Rows[1, misValue].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

            //#endregion CREATE REPORT

            //#region CREATE SUMMARY

            ////Get the second Worksheet of the active WorkBook.
            //// xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            //xlWorkBook.Worksheets.Add(Type.Missing, (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2), misValue, misValue);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
            //xlWorkSheet.Name = "Summary";

            ////Header
            //xlWorkSheet.Range["B2", "H2"].Merge(false);
            //xlWorkSheet.Range["B2", "H2"].Value = "ΓΕΝΙΚΟ ΣΥΝΟΛΟ";
            //xlWorkSheet.Range["B2", "H2"].RowHeight = 26;
            //xlWorkSheet.Range["B2", "H2"].Font.Bold = true;
            //xlWorkSheet.Range["B2", "H2"].Font.Size = 18;
            //xlWorkSheet.Range["B2", "H2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Range["B2", "H2"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //xlWorkSheet.Range["B2", "H2"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            //xlWorkSheet.Range["B2", "H2"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["B2", "H2"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            ////Description
            //xlWorkSheet.Range["B3", "C3"].Merge(false);
            //xlWorkSheet.Range["B3", "C3"].Value = "ΠΕΡΙΓΡΑΦΗ";
            //xlWorkSheet.Range["B3", "C3"].RowHeight = 18;
            //xlWorkSheet.Range["B3", "C3"].Font.Bold = true;
            //xlWorkSheet.Range["B3", "C3"].Font.Size = 12;
            //xlWorkSheet.Range["B3", "C3"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //xlWorkSheet.Range["B3", "C3"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //xlWorkSheet.Range["B3", "C3"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            //xlWorkSheet.Range["B3", "C3"].Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //xlWorkSheet.Range["B3", "C3"].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            #endregion CREATE SUMMARY

            //Save Excel file.
            xlWorkBook.SaveAs(excelFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            //Release Excel file.
            // Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Excel file created!!");

        }

        private void menuExportImage_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog _saveFileDialog = new SaveFileDialog();
            if (_saveFileDialog.ShowDialog() == true)
            {
                View1.Export(_saveFileDialog.FileName);
            }
        }

        private void menuRemoveDuplicates_Click(object sender, RoutedEventArgs e)
        {
            var distinct = lstHeadLossCalcs.Distinct<HeadLossCalc>(new ItemEqualityComparer()).ToList<HeadLossCalc>();
            lstHeadLossCalcs.Clear();
            foreach (var item in distinct)
            {
                lstHeadLossCalcs.Add(item);
            }
        }

        private void tabControl1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabControl tb = sender as TabControl;

            if (((TabItem)tb.SelectedItem).Header.ToString() == "Components")
            {
                GroupElements();
            }
        }

        private void dgvBOM_SelectionChanged(object sender, Xceed.Wpf.DataGrid.DataGridSelectionChangedEventArgs e)
        {
            HeadLossCalc selectedItem = (HeadLossCalc)e.SelectionInfos[0].DataGridContext.CurrentItem;
            if (selectedItem != null)
            {
                double totalW = 0;
                double totalV = 0;
                double totalL = 0;
                var selection = dgvBOM.SelectedItems;
                foreach (HeadLossCalc hlc in selection)
                {
                    totalL = totalL + hlc.L;
                    totalV = totalV + hlc.Vm;
                    totalW = totalW + hlc.W;
                }
                tbL2.Text = totalL.ToString("N3");
                tbV2.Text = totalV.ToString("N3");
                tbW2.Text = totalW.ToString("N2");
            }
        }

        private void dgvSummary_SelectionChanged(object sender, Xceed.Wpf.DataGrid.DataGridSelectionChangedEventArgs e)
        {

        }

        private void menuUpdateReport_Click(object sender, RoutedEventArgs e)
        {
            GroupElements();
        }

        private void menuGroupAll_Click(object sender, RoutedEventArgs e)
        {
            GroupElements();
        }

        private void menuGroupLines_Click(object sender, RoutedEventArgs e)
        {
            GroupElements(false);
        }

        private void GroupElements(bool groupAll = true)
        {
            dgvBOM.ItemsSource = lstHeadLossCalcs;

            if (!groupAll)
            {
                dgvSummary.Columns[0].Visible = true;

                var r = lstHeadLossCalcs
                     .GroupBy(x => new { x.Line, x.ElementType, x.d, x.d1, x.r, x.a, x.Notes })
                     .Select(y => new
                     {
                         Line = y.Key.Line,
                         Element = y.Key.ElementType,
                         Diameter = y.Key.d,
                         D1 = y.Key.d1,
                         Length = y.Sum(x => x.L),
                         Radious = y.Key.r,
                         Angle = y.Key.a,
                         Weight = y.Sum(x => x.W),
                         Volume = y.Sum(x => x.Vm),
                         Count = y.Count(),
                         Notes = y.Key.Notes
                     })
                     .OrderBy(x => x.Line).ThenBy(x => x.Element).ThenBy(x => x.Diameter).ThenBy(x => x.D1).ThenBy(x => x.Radious).ThenBy(x => x.Angle)
                     .ToList();
                dgvSummary.ItemsSource = r;
            }
            else
            {

                dgvSummary.Columns[0].Visible = false;

                var r = lstHeadLossCalcs
                                     .GroupBy(x => new { x.ElementType, x.d, x.d1, x.r, x.a, x.Notes })
                                     .Select(y => new
                                     {
                                         Line = "",
                                         Element = y.Key.ElementType,
                                         Diameter = y.Key.d,
                                         D1 = y.Key.d1,
                                         Length = y.Sum(x => x.L),
                                         Radious = y.Key.r,
                                         Angle = y.Key.a,
                                         Weight = y.Sum(x => x.W),
                                         Volume = y.Sum(x => x.Vm),
                                         Count = y.Count(),
                                         Notes = y.Key.Notes
                                     })
                                     .OrderBy(x => x.Element).ThenBy(x => x.Diameter).ThenBy(x => x.D1).ThenBy(x => x.Radious).ThenBy(x => x.Angle)
                                     .ToList();
                dgvSummary.ItemsSource = r;
            }

        }

    }
}
