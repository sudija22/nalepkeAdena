using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using SpreadsheetLight;
using System.Collections;


namespace nalepkeAdena
{
    public partial class Form1 : Form
    {


        string modelBed;
        string sizeXBed;
        string sizeYBed;
        string quantityBed;
        double bedRowNum;
        string bedOrderNumLocal;
        string bedOrderNumber;
        string bedDeliveryCompany;
        string bedRif;
        string bedDescription;



        string datoteka = null;
        string ordineFrame;
        string stickerDate;
        string deliveyCompanyFrame;
        string vVFrame;
        string rifFrame;
        string motorFrame;
        string mountTypeFrame;
        string legsFrame;
        string italCodeFrame;
        string personalizationFrame;
        string packingFrame;
        string descriptionFrame;
        List < String > adsFrame = new List<String>();
        string firstAdFrame;
        string secondAdFrame;
        string ad1="";
        string ad2 = "";
        string ad3 = "";
        string modelFrame;
        string typeFrame; //1,2,3
        string sizeXFrame;
        string sizeYFrame;
        int piecesFrame;
        string narocilo;
        string dodatek1;
        string lastnost1;
        string oznakaModel;
        string dodatek4; // new model, 60x40
        string kraj;
        string rifInKraj;
        string guma;
        string vrsticaCheck = "";
        string lastnost2;
        string dodatek2;
        string dodatek3;
        string cbText1 = "CON FORATURE E CARATTERISTICHE";
        string cbText2 = "X CONTENITORE ERGOGREEN";
        string ccText1 = "CON FORI X MECCANISMO CONFORT";
        string clText1 = "CON FORATURE E CARATTERISTICHE";
        string clText2 = "X LETTO LIFT";
        string lfText1 = "CON FORATURE X SISTEMA LIFTER";
        string goText1 = "CON VELCRO X GONNELLINO";
        string auText1 = "FORI X SPONDE AUXILIA";
        string motoriNapis = "MOTORE MONTATO";
        string[] listaOblecene = { "MICHELLE", "ALEXIA","ASIA","DREAM", "SOMMIER", "FREESTYLE"};

        public bool checkSpecial(string model)
        {
            for(int i=0; i<listaOblecene.Length; i++)
            {
                if (model == listaOblecene[i])
                {
                    return true;
                }
            }
            return false;
        }

        public Form1()
        {
            InitializeComponent();
        }


        private void fileSystemWatcher1_Changed(object sender, System.IO.FileSystemEventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)//izberi datoteko
        {
            int size = -1;
            var fileContent = string.Empty;
            var filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK) // Test result.
                {
                    datoteka = openFileDialog.FileName;
                    try
                    {
                        {
                           
                            //Get the path of specified file
                            filePath = datoteka;
                            //Read the contents of the file into a stream
                            var fileStream = openFileDialog.OpenFile();

                            using (StreamReader reader = new StreamReader(fileStream))
                            {
                                fileContent = reader.ReadToEnd();
                            }
                        }

                        string text = File.ReadAllText(datoteka);
                        size = text.Length;
                    }
                    catch (IOException napaka)
                    {
                    }
                }

            }

            
            
        }
        private Boolean checkFileFormat(object sender, EventArgs e, string datoteka)
        {
            try {
                if (datoteka == null)
                {
                    MessageBox.Show("Najprej izberi datoteko");
                    return false;
                }
                string formatCheck = datoteka.Substring((datoteka.Length - 4), 4);
                if (formatCheck != "xlsx") {
                    MessageBox.Show("Format ni pravilen. Pravilen format je '.xlsx'");
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch(IOException er) {
                return false;
            }
        }

        private void potrditev_Click(object sender, EventArgs e)
        {
            if (checkFileFormat(sender, e, datoteka)) { 
             
                SLDocument fileNarocila = new SLDocument(datoteka); //open order file

                string pathPredloga = "../";
                string kocnoPredlogaPath = pathPredloga + "\\template.xlsx";
                
                
                SLWorksheetStatistics stats = fileNarocila.GetWorksheetStatistics(); // stats for order file, to get last row

                
                
                
                SLDocument frameLabelFinalFile = new SLDocument("template.xlsx"); //open order file
                
                int stevec = 2;
                
                for (int i = 3; i <= stats.NumberOfRows; i++)
                {
                    ordineFrame = fileNarocila.GetCellValueAsString(i, 4);
                    rifFrame = fileNarocila.GetCellValueAsString(i, 16);
                    vVFrame = fileNarocila.GetCellValueAsString(i, 8);
                    deliveyCompanyFrame = fileNarocila.GetCellValueAsString(i, 15);
                    DateTime today = DateTime.Today;
                    string[] collection = today.ToString("d").Split('.'); 
                    stickerDate = (String.Format("{0}{1}", collection[0], collection[1].Trim())).Trim();

                    modelFrame = fileNarocila.GetCellValueAsString(i, 6);
                    if(modelFrame=="EVO   SATURNO")
                    {
                        modelFrame = "SATURNO";
                    }
                    if (modelFrame == "EVO   PLUTONE")
                    {
                        modelFrame = "PLUTONE";
                    }
                    if (modelFrame == "EVO   NETUNO")
                    {
                        modelFrame = "NETTUNO";
                    }
                    if (modelFrame == "EVO  PT  NETTUNO")
                    {
                        modelFrame = "NETTUNO";
                    }
                    if (modelFrame == "EVO  PT  PLUTONE")
                    {
                        modelFrame = "PLUTONE";
                    }
                    
                    typeFrame = fileNarocila.GetCellValueAsString(i, 7);
                    sizeXFrame = fileNarocila.GetCellValueAsString(i, 9);
                    sizeYFrame = fileNarocila.GetCellValueAsString(i, 10);
                    piecesFrame = fileNarocila.GetCellValueAsInt32(i, 22);
                    motorFrame = fileNarocila.GetCellValueAsString(i, 20);
                    mountTypeFrame = fileNarocila.GetCellValueAsString(i, 11);
                    packingFrame   = fileNarocila.GetCellValueAsString(i, 12);
                    italCodeFrame = fileNarocila.GetCellValueAsString(i, 1);
                    personalizationFrame = fileNarocila.GetCellValueAsString(i, 5);
                    legsFrame = fileNarocila.GetCellValueAsString(i, 19);
                    descriptionFrame = fileNarocila.GetCellValueAsString(i, 17);
                    firstAdFrame = fileNarocila.GetCellValueAsString(i, 13);
                    secondAdFrame = fileNarocila.GetCellValueAsString(i, 14);

                    // Add some text to file    
                   adsFrame.Clear();
                    if (packingFrame == "C")
                    {
                        adsFrame.Add("CONCARTONE");
                    }// c concartone
                    
                    if(motorFrame=="T2" || motorFrame=="T3" || motorFrame=="T6" || motorFrame == "T56")
                    {
                        adsFrame.Add("MOTORE MONTATA");
                    }
                    if (legsFrame != "" || legsFrame != null)
                    {
                        adsFrame.Add(legsFrame);
                    }
                    if (descriptionFrame != "" || descriptionFrame != null)
                    {
                        adsFrame.Add(descriptionFrame);
                    } //napoljnen seznam ads
                    if (firstAdFrame != "" || firstAdFrame != null)
                    {
                        adsFrame.Add(firstAdFrame);
                    } //napoljnen seznam ads
                    if (secondAdFrame != "" || secondAdFrame != null)
                    {
                        adsFrame.Add(secondAdFrame);
                    } //napoljnen seznam ads


                    for (int k = 0; k < adsFrame.Count; k++)
                    {

                        if (ad1.Length+adsFrame[k].Length <= 16)
                        {
                            ad1 = ad1 + adsFrame[k] + " ";
                        }
                        else
                        {
                            if (ad2.Length + adsFrame[k].Length <= 16)
                            {
                                ad2 = ad2 + adsFrame[k] + " ";
                            }
                            else
                            {
                                if (ad3.Length + adsFrame[k].Length <= 28)
                                {
                                    ad3 = ad3 + adsFrame[k] + " ";

                                }
                            }
                        }

                    }

                    SLStyle odebeljeno = frameLabelFinalFile.CreateStyle();
                    odebeljeno.Font.Bold = true;
                    odebeljeno.Font.FontName = "Arial CE";
                    for (int j = 1; j <= piecesFrame; j++)
                    {
                        //prva vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 1, "ORDINE:");
                        frameLabelFinalFile.SetCellValue(stevec, 3, ordineFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7, stickerDate);
                        frameLabelFinalFile.SetCellValue(stevec, 1 + 13, "ORDINE:");
                        frameLabelFinalFile.SetCellValue(stevec, 3 + 13, ordineFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7 + 13, stickerDate);
                        stevec++;

                        //druga vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 1+ 13, "RIF:");
                        frameLabelFinalFile.SetCellValue(stevec, 2, deliveyCompanyFrame+rifFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 1 + 13, "RIF:");
                        frameLabelFinalFile.SetCellValue(stevec, 2 + 13, deliveyCompanyFrame + rifFrame);
                        stevec++;


                        //tretja vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 7, mountTypeFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7 + 13, mountTypeFrame);
                        if (mountTypeFrame == "CB")
                        {
                            frameLabelFinalFile.SetCellValue(stevec, 1, "CON FORI X MECCANISMO CONFORT");
                            frameLabelFinalFile.SetCellValue(stevec, 1 + 13,"CON FORI X MECCANISMO CONFORT");
                        }
                        stevec++;

                        //cetrta vrstica 
                        
                        stevec++;

                        //peta vrstica
                        SLStyle fontMereLength = frameLabelFinalFile.CreateStyle();
                        fontMereLength.Font.FontName = "Arial CE";
                        if (modelFrame.Length > 9)
                        {
                            fontMereLength.Font.FontSize = 14;
                            frameLabelFinalFile.SetCellValue(stevec, 1, modelFrame);
                            frameLabelFinalFile.SetCellStyle(stevec, 1, fontMereLength);
                            frameLabelFinalFile.SetCellValue(stevec, 1 + 13, modelFrame);
                            frameLabelFinalFile.SetCellStyle(stevec, 1+13, fontMereLength);
                            fontMereLength.Font.FontSize = 16;
                        }
                        else
                        {
                            frameLabelFinalFile.SetCellValue(stevec, 1, modelFrame);
                            frameLabelFinalFile.SetCellValue(stevec, 1 + 13, modelFrame);
                        }
                        frameLabelFinalFile.SetCellValue(stevec, 7, typeFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 6, vVFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7 + 13, typeFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 6 + 13, vVFrame);
                        stevec++;

                        //sesta vrstica
                        stevec++;

                        //sedma vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 1, sizeXFrame + "X" + sizeYFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 1 + 13, sizeXFrame + "X" + sizeYFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 4, ad1);
                        frameLabelFinalFile.SetCellValue(stevec, 4 + 13, ad1);
                        //osma vrstica
                        stevec++;
                        frameLabelFinalFile.SetCellValue(stevec, 4, ad2);
                        frameLabelFinalFile.SetCellValue(stevec, 4 + 13, ad2);
                        //deveta vrstica
                        stevec++;
                        frameLabelFinalFile.SetCellValue(stevec, 1, ad3);
                        frameLabelFinalFile.SetCellValue(stevec, 1 + 13, ad3);
                        frameLabelFinalFile.SetCellStyle(stevec, 1, odebeljeno);
                        frameLabelFinalFile.SetCellStyle(stevec, 1+13, odebeljeno);

                        //to bos izbrisal drugic
                        stevec = stevec + 3;
                        
                    }
                    ad1 = ad2 = ad3 = "";

                }
                
                DateTime thisDay = DateTime.Today;
                string path = "./"; //get current path
                string shrani = path + "\\" + thisDay.ToString("d") + "NALEPKE.xlsx"; // format save name of file to save on user destop
                MessageBox.Show(shrani);
                frameLabelFinalFile.SaveAs(shrani); //save sticker file

                

                    fileNarocila.CloseWithoutSaving(); //close order file
                    MessageBox.Show("Nalepke so kreirane."); //messsage shot for successful sticker create



            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void createBarcode_Click(object sender, EventArgs e)
        {
            string barCodeTest = barcodeText.Text;
            try
            {
                Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                pictureBox2.Image = brCode.Draw(barCodeTest, 50); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali

            }
            catch
            {

            }
        }

        private void btnBedList_Click(object sender, EventArgs e)
        {
            var fileBedContent = string.Empty;
            var fileBedPath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK) // Test result.
                {
                    datoteka = openFileDialog.FileName;
                    try
                    {
                        {

                            //Get the path of specified file
                            fileBedContent = datoteka;
                            //Read the contents of the file into a stream
                            var fileStream = openFileDialog.OpenFile();

                            using (StreamReader reader = new StreamReader(fileStream))
                            {
                                fileBedContent = reader.ReadToEnd();
                            }
                        }

                        string text = File.ReadAllText(datoteka);
                    }
                    catch (IOException napaka)
                    {

                    }
                }

            }
        
        }

        private void btnCreateLabelBed_Click(object sender, EventArgs e)
        {
            if (checkFileFormat(sender, e, datoteka))
            {
                SLDocument fileBedList = new SLDocument(datoteka); //open order file
                SLWorksheetStatistics stats = fileBedList.GetWorksheetStatistics(); // stats for order file, to get last row
                int lastRowBedIndex = stats.NumberOfRows;


                for (int i = 4 ; i<lastRowBedIndex; i++)
                {
                    bedRowNum = fileBedList.GetCellValueAsDouble(i, 1);
                    bedOrderNumLocal = fileBedList.GetCellValueAsString(i, 2);
                    bedOrderNumber = fileBedList.GetCellValueAsString(i, 3);
                    modelBed = fileBedList.GetCellValueAsString(i, 5);
                    sizeXBed = fileBedList.GetCellValueAsString(i, 8);
                    sizeYBed = fileBedList.GetCellValueAsString(i, 9);
                    bedDeliveryCompany = fileBedList.GetCellValueAsString(i, 14);
                    bedRif = fileBedList.GetCellValueAsString(i, 15);
                    bedDescription = fileBedList.GetCellValueAsString(i, 16);
                    quantityBed = fileBedList.GetCellValueAsString(i, 17);


                    Console.WriteLine(modelBed+bedOrderNumber);
                }




                

                string pathPredloga = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string kocnoPredlogaPath = pathPredloga + "\\nalepkeProgram\\nalepke.xlsx";
                
                MessageBox.Show("Nalepke so kreirane."); //messsage shot for successful sticker create



            }
        }
    }
}

//time 18,5 h +2
// todo list:
// bonsai salus string length   
//AUGO vse v isto polje. 
// kaj je z innovo 3/2 ???
//no pistoni v rifu ??        

