using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExtensionMethods;
using System.IO;
using System.Collections.ObjectModel;
using System.Drawing.Imaging;

namespace CBSys
{

    public partial class ThisWorkbook
    {
        public string[] Renglones = Properties.Resources.Content.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
        
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            ObservableCollection<string> listClientes= GetListClientes();

            Excel.Workbook wb = Application.Workbooks["CBSys.xlsx"];
            Excel.Worksheet Sheet = wb.Sheets["Clientes"];
            Excel.Range rngCliente = Sheet.Cells[1, 1];
            for (int i = 0; i < listClientes.Count; i++)
            {
                rngCliente.Offset[i, 0].Value = listClientes[i].ToString();
            }
            Neodynamic.SDK.Barcode.BarcodeProfessional.LicenseOwner = "Grupo Empresarial CT SA de CV-Ultimate Edition-OEM Developer License";
            Neodynamic.SDK.Barcode.BarcodeProfessional.LicenseKey = "CWELTAJF8DTZP4RQEB6Y9RFPDH6EHE8J29PYSKF5B9LPZQHSMZJQ";
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {

        }

        private ObservableCollection<string> GetListClientes()
        {
            ObservableCollection<string> listClientes = new ObservableCollection<string>();
            for (int i = 0; i < Renglones.Length; i++)
            {
                if (!listClientes.Contains(Renglones[i].GetClient()))
                {
                    listClientes.Add(Renglones[i].GetClient());
                }
            }
            return listClientes;
        }

        private void SChange()
        {

        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.SheetSelectionChange += new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetSelectionChangeEventHandler(this.SChange);
            this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(this.ThisWorkbook_Shutdown);

        }

        #endregion

        private void SChange(object Sh, Excel.Range Target)
        {
            if(Target.Value==null)
            {
                Excel.Workbook wb = Application.Workbooks["CBSys.xlsx"];
                Excel.Worksheet Sheet = wb.Sheets["Clientes"];
                Excel.Range rngModelo = Sheet.Cells[1, 2];
                rngModelo.EntireColumn.ClearContents();
            }
            if (Target.Column.ToString() == "1" && Target.Value!=null)
            {
                Target.Offset[0,1].EntireColumn.ClearContents();
                Excel.Workbook wb = Application.Workbooks["CBSys.xlsx"];
                Excel.Worksheet Sheet = wb.Sheets["Clientes"];
                Excel.Range rngModelo = Sheet.Cells[1, 2];
                int count = 0;
                for (int i=0;i<Renglones.Length;i++)
                {
                    if(Target.Value==Renglones[i].GetClient())
                    {
                        rngModelo.Offset[count, 0].Value = Renglones[i].GetModel();
                        count = count + 1;
                    }
                }
            }
            string gs = "\x1d";
            if (Target.Column.ToString()=="2" && Target.Value!=null)
            {
                Neodynamic.SDK.Barcode.BarcodeProfessional brcd = new Neodynamic.SDK.Barcode.BarcodeProfessional();
                string modelo = "";
                string cliente = "";
                string codigo = "";
                for (int i=0;i<Renglones.Length;i++)
                {
                    if(Renglones[i].GetModel()==Target.Value)
                    {
                        codigo = "~6P"+Renglones[i].GetModel()+gs+"S"+DateTime.Now.ToString("ddMMyyyy")+gs;
                        modelo = Renglones[i].GetModel();
                        cliente = Renglones[i].GetClient();
                    }
                }
                string pathimage = "J:\\Edge\\" + DateTime.Today.ToString("ddMMyyyy")+"\\"+ cliente + "\\" + modelo + "\\" + modelo + ".png";
                brcd.Symbology = Neodynamic.SDK.Barcode.Symbology.DataMatrix;
                Directory.CreateDirectory("J:\\Edge\\" + DateTime.Today.ToString("ddMMyyyy")+"\\"+ cliente.ToString() + "\\" + modelo.ToString());
                brcd.DataMatrixProcessTilde = true;
                brcd.Code = codigo;
                brcd.BorderWidth = 0;
                brcd.BorderWidth = 0;
                brcd.BottomMargin = 0;
                brcd.TopMargin = 0;
                brcd.Width = 1;


                //MessageBox.Show("brcd.BarWidthAdjustment: " + brcd.BarWidthAdjustment.ToString()+"\n"+"brcd.BottomMargin: "+brcd.BottomMargin+"brcd.DisplayLightMarginIndicator: "+brcd.DisplayLightMarginIndicator + "brcd.Image.VerticalResolution: "+brcd.Image.VerticalResolution+ "brcd.BackColor.ToKnownColor().ToString(): " + brcd.BackColor.ToKnownColor().ToString());
                brcd.QuietZoneWidth = 0;
                brcd.Width = 10;
                System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(brcd.GetBarcodeImage());

				Neodynamic.SDK.Barcode.BarcodeProfessional.LicenseOwner = "Grupo Empresarial CT SA de CV-Ultimate Edition-OEM Developer License";
				Neodynamic.SDK.Barcode.BarcodeProfessional.LicenseKey = "CWELTAJF8DTZP4RQEB6Y9RFPDH6EHE8J29PYSKF5B9LPZQHSMZJQ";

				ImageCodecInfo codecinfo = GetEncoder(ImageFormat.Png);
                System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;
                EncoderParameters myencoderparameters = new EncoderParameters(1);
                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 1000L);
                myencoderparameters.Param[0] = myEncoderParameter;
                bmp.Save(pathimage,codecinfo,myencoderparameters);
            }
        }
        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
			Neodynamic.SDK.Barcode.BarcodeProfessional.LicenseOwner = "Grupo Empresarial CT SA de CV-Ultimate Edition-OEM Developer License";
			Neodynamic.SDK.Barcode.BarcodeProfessional.LicenseKey = "CWELTAJF8DTZP4RQEB6Y9RFPDH6EHE8J29PYSKF5B9LPZQHSMZJQ";

			ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
                if (codec.FormatID == format.Guid)
                    return codec;
            return null;
        }
    }
}
