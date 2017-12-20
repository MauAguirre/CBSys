using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using Neodynamic.SDK;
using System.Drawing;
using System.Collections.ObjectModel;
using ExtensionMethods;
using Microsoft.Office.Interop.Excel;



namespace CBSys
{
    public partial class RibCBSys
    {
        public static List<string> formatosDataMatrix = new List<string>();
        public string[] Renglones = Properties.Resources.Content.Split(new string[] { Environment.NewLine },StringSplitOptions.RemoveEmptyEntries);
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            List<string> symbologies = new List<string>();
            symbologies.Add("DataMatrix");
            formatosDataMatrix.Add("Formato 06");
            RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item.Label = symbologies[0];
            drSymbology.Items.Add(item);
            SetFormat();
        }
        private void drSymbology_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            SetFormat();
        }
        private void SetFormat()
        {
            if (drSymbology.SelectedItem.Label == "DataMatrix")
            {
                drFormato.Items.Clear();
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = formatosDataMatrix[0];
                drFormato.Items.Add(item);
            }
        }



        private void btnRun_Click(object sender, RibbonControlEventArgs e)
        {
            if(drSymbology.SelectedItem.Label=="DataMatrix")
            {

            }
        }
    }
}
namespace ExtensionMethods
{
    public static class MyExtensions
    {
        public static string GetID(this string str)
        {
            char[] delimiters = { ',' };
            string[] arrstr = str.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            return arrstr[0];
        }
        public static string GetClient(this string str)
        {
            char[] delimiters = { ',' };
            string[] arrstr = str.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            return arrstr[1];
        }
        public static string GetModel(this string str)
        {
            char[] delimiters = { ',' };
            string[] arrstr = str.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            return arrstr[2];
        }
        public static string GetFormat(this string str)
        {
            char[] delimiters = { ',' };
            string[] arrstr = str.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            return arrstr[3];
        }
        public static string GetContent(this string str)
        {
            char[] delimiters = { ',' };
            string[] arrstr = str.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            arrstr[4] = arrstr[4].Replace("!", DateTime.Now.ToString("ddMMyyyy"));
            return arrstr[4];
        }
    }
}
