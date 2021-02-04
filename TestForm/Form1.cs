using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestForm
{
    public partial class Form1 : Form
    {
        string pathAsks="";
        string pathRKK="";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pathAsks = openFileDialog.FileName;
                label1.Text = pathAsks;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pathRKK = openFileDialog.FileName;
                label2.Text = pathRKK;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Stopwatch myStopwatch = new System.Diagnostics.Stopwatch();
            myStopwatch.Start(); 

            if (pathAsks != "" & pathRKK != "")
            {
                Dictionary<string, int[]> resultDict = new Dictionary<string, int[]>();
                resultDict=Calculate(pathRKK, pathAsks);

                int i = 1;
                foreach (KeyValuePair<string, int[]> tmp in resultDict)
                    Console.WriteLine("Key = {0}, Value = {1}, Value = {2}, Value = {3}, i = {4}",
                                        tmp.Key, tmp.Value[0], tmp.Value[1], tmp.Value[2], i++); 
                myStopwatch.Stop();
                Form2 f = new Form2(resultDict, myStopwatch.Elapsed);
                f.Show();
            }
            else
                MessageBox.Show("Не выбраны файлы!");
        }
        static Dictionary<string, int[]> Calculate(string pathRKK, string pathAsks)
        {
            Dictionary<string, int> rkk = new Dictionary<string, int>();
            Dictionary<string, int> asks = new Dictionary<string, int>();
            rkk = ReadFile(pathRKK);
            asks = ReadFile(pathAsks);

            Dictionary<string, int[]> unionDict = new Dictionary<string, int[]>();
            foreach (KeyValuePair<string, int> tmp in rkk)
            {
                unionDict.Add(tmp.Key, new int[3]);
                unionDict[tmp.Key][0] = tmp.Value;
                unionDict[tmp.Key][1] = 0;
                unionDict[tmp.Key][2] = tmp.Value;
            }
            foreach (KeyValuePair<string, int> tmp in asks)
            {
                if (!unionDict.ContainsKey(tmp.Key))
                {
                    unionDict.Add(tmp.Key, new int[3]);
                    unionDict[tmp.Key][0] = 0;
                    unionDict[tmp.Key][1] = tmp.Value;
                    unionDict[tmp.Key][2] = tmp.Value;
                }
                else
                {
                    unionDict[tmp.Key][1] = tmp.Value;
                    unionDict[tmp.Key][2] = tmp.Value + unionDict[tmp.Key][0];
                }
            }
            var sortedDict = (from unionD in unionDict
                             orderby unionD.Value[2] descending,
                                     unionD.Value[1] descending,
                                     unionD.Value[0] descending,
                                     unionD.Key descending
                             select unionD).ToDictionary(unionD => unionD.Key, unionD => unionD.Value);

            return sortedDict;
        }

        static Dictionary<string, int> ReadFile(string path)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            using (StreamReader sr = new StreamReader(path, System.Text.Encoding.Default))
            {
                string line;
                string[] manager;
                string[] executor;
                while ((line = sr.ReadLine()) != null)
                {
                    manager = line.Split('\t');
                    if (manager[0].Contains("Климов"))
                    {
                        executor = manager[1].Split(';');
                        manager[0] = executor[0].Replace("(Отв.)", "").Trim();
                    }
                    else
                        manager[0] = ExctraxtIni(manager[0]);

                    if (!dict.ContainsKey(manager[0]))
                    {
                        dict.Add(manager[0], 1);
                    }
                    else
                        dict[manager[0]] += 1;
                }
            }
            return dict;
        }

        static string ExctraxtIni(string s)
        {
            var inits = Regex.Match(s, @"(\w+)\s+(\w+)\s+(\w+)").Groups;
            return string.Format("{0} {1}.{2}.", inits[1], inits[2].Value[0], inits[3].Value[0]);
        }


    }
}
