using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

using MyExcel;

namespace ZeroTouchTekla
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void OnSetExcelClick(object sender, EventArgs e)
        {
            string path = ExcelUtility.BrowserPath();
            if (ExcelUtility.isValidExcelPath(path))
            {
                Program.ExcelPath = path;
                ExcelNameLabel.Text = Path.GetFileName(path);
                ExcelUtility excelUtility = new ExcelUtility(path);
                List<string> sheetNames = excelUtility.SheetList();
                excelUtility.SaveAndDispose();
                SheetComboBox.Items.Clear();
                SheetComboBox.Items.AddRange(sheetNames.ToArray());
            }
        }

        private void OnSelectedSheetChanged(object sender, EventArgs e)
        {
            Program.SheetName = SheetComboBox.SelectedItem.ToString();
        }

        private void OnLoadButtonClick(object sender, EventArgs e)
        {

            if (ExcelUtility.isValidExcelPath(Program.ExcelPath))
            {
                ExcelUtility excelUtility = new ExcelUtility(Program.ExcelPath);

                Program.ExcelDictionary = excelUtility.ReadExcelAndCreateDictionary(3, 4, Program.SheetName);
                excelUtility.SaveAndDispose();
                /*
                catch
                {
                    MessageBox.Show("Cannot load excel", "Error");

                }
                */
            }
            else
            {
                MessageBox.Show("Excel path is not valid", "Error");
            }

        }

        private void OnFootingRebarClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.FTG);
        }

        private void OnCopySpacingClick(object sender, EventArgs e)
        {
            Utility.CopyGuideLine();
        }

        private void OnRecreateRebarButtonClick(object sender, EventArgs e)
        {
            RebarCreator.RecreateRebar();
        }

        private void OnRTWButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.RTW);
        }

        private void OnTestButtonClick(object sender, EventArgs e)
        {
            RebarCreator.Test();

        }

        private void OnCollumnButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.CLMN);
        }

        private void OnDoubleRTWButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.DRTW);
        }

        private void OnRetainingWallStepButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.RTWS);
        }

        private void OnAbutmentButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.ABT);
        }

        private void OnWingButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForComponent(Element.ProfileType.WING);
        }

        private void OnTripledAbutmentButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.TABT);
        }

        private void OnAPSButtonClick(object sender, EventArgs e)
        {
            RebarCreator.CreateForPart(Element.ProfileType.APS);
        }
    }
}
