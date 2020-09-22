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
using SpreadsheetLight.Drawing;

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
        string barCodeTest = "";
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
        List<String> adsFrame = new List<String>();
        List<Image> images = new List<Image>();
        int barCodeIndex=11;
        int barCounter = 0;
        string firstAdFrame;
        string secondAdFrame;
        string ad1 = "";
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
        string[] listaOblecene = { "MICHELLE", "ALEXIA", "ASIA", "DREAM", "SOMMIER", "FREESTYLE" };

        public bool checkSpecial(string model)
        {
            for (int i = 0; i < listaOblecene.Length; i++)
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
                            Console.WriteLine(datoteka);
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
            try
            {
                if (datoteka == null)
                {
                    MessageBox.Show("Najprej izberi datoteko");
                    return false;
                }
                string formatCheck = datoteka.Substring((datoteka.Length - 4), 4);
                if (formatCheck != "xlsx")
                {
                    MessageBox.Show("Format ni pravilen. Pravilen format je '.xlsx'");
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (IOException er)
            {
                return false;
            }
        }

        private void potrditev_Click(object sender, EventArgs e)
        {
            if (checkFileFormat(sender, e, datoteka))
            {

                SLDocument fileNarocila = new SLDocument(datoteka); //open order file
                Console.WriteLine("tukaj");
                Console.WriteLine(datoteka);
                Console.WriteLine(fileNarocila.GetCellValueAsString(1, 1));
                //SLDocument fileNalepke = new SLDocument("template.xlsx");// open template file
                string pathPredloga = "../";
                string kocnoPredlogaPath = pathPredloga + "\\template.xlsx";

                // SLDocument fileNalepke = new SLDocument(kocnoPredlogaPath);
                //SLDocument fileNalepkeOblecene = new SLDocument("predlogaOblecene.xlsx");
                //     SLStyle fontMereLength = fileNalepke.CreateStyle();
                //    SLStyle fontModelLength1 = fileNalepke.CreateStyle();
                //    SLStyle odebeljeno = fileNalepke.CreateStyle();

                /*    fontMereLength.Font.FontSize = 14; // change if string lengt of size is too long 9 chars.
                    fontModelLength1.Font.FontSize = 16;
                    fontMereLength.Font.FontName = "Arial CE";
                    fontModelLength1.Font.FontName = "Arial CE";
                    odebeljeno.Font.Bold = true;
                    odebeljeno.Font.FontName = "Arial CE";

                    SLStyle fontModelLength2 = fileNalepke.CreateStyle(); 
                    fontModelLength2.Font.FontSize = 13.5;
                    fontModelLength2.Font.FontName = "Arial CE"; 

                Random r = new Random(); //kaj je to ?

                    int kazalec = 2;

                    string datum = fileNarocila.GetCellValueAsString("F1"); // get date and format it 
                    string[] splitanDatum = datum.Split('.');
                    string formatiranDatum = "";
                    if (splitanDatum[0].Length != 2)
                    {
                        formatiranDatum += "0" + splitanDatum[1];
                    }
                    else
                    {
                        formatiranDatum += splitanDatum[1];
                    }
                    string leto = splitanDatum[2].Substring((splitanDatum[2].Length - 2), 2);
                    formatiranDatum += leto; */


                //MessageBox.Show(steviloStrani.ToString());
                SLWorksheetStatistics stats = fileNarocila.GetWorksheetStatistics(); // stats for order file, to get last row



                //SLDocument frameLabelFinalFile = new SLDocument("C:\\Users\\tomaz\\Desktop\\Novica.xlsx"); //open order file

                SLDocument frameLabelFinalFile = new SLDocument("template.xlsx"); //open order file
                DateTime thisDay = DateTime.Today;
                string path = "./"; //get current path
                string shrani = path + "\\" + thisDay.ToString("d") + "NALEPKE.xlsx";
                /*  DateTime thisDay = DateTime.Today;
                  Console.WriteLine(thisDay.ToString("d"));
                  string path = "./"; //get current path
                  string shrani = path + "\\" + thisDay.ToString("d") + "NALEPKE.xlsx"; // format save name of file to save on user destop
                  MessageBox.Show(shrani);
                  frameLabelFinalFile.SaveAs(shrani); //save sticker file */

                int stevec = 2;

                for (int i = 3; i <= stats.NumberOfRows; i++)
                {
                    Console.WriteLine(stats.NumberOfRows);
                    Console.WriteLine(i);
                    string mess = "halooo";
                    ordineFrame = fileNarocila.GetCellValueAsString(i, 4);
                    rifFrame = fileNarocila.GetCellValueAsString(i, 16);
                    vVFrame = fileNarocila.GetCellValueAsString(i, 8);
                    deliveyCompanyFrame = fileNarocila.GetCellValueAsString(i, 15);
                    DateTime today = DateTime.Today;
                    string[] collection = today.ToString("d").Split('.');
                    stickerDate = (String.Format("{0}{1}", collection[0], collection[1].Trim())).Trim();
                    Console.WriteLine(stickerDate);

                    modelFrame = fileNarocila.GetCellValueAsString(i, 6);
                    if (modelFrame == "EVO   SATURNO")
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
                    packingFrame = fileNarocila.GetCellValueAsString(i, 12);
                    italCodeFrame = fileNarocila.GetCellValueAsString(i, 1);
                    personalizationFrame = fileNarocila.GetCellValueAsString(i, 5);
                    legsFrame = fileNarocila.GetCellValueAsString(i, 19);
                    descriptionFrame = fileNarocila.GetCellValueAsString(i, 17);
                    firstAdFrame = fileNarocila.GetCellValueAsString(i, 13);
                    secondAdFrame = fileNarocila.GetCellValueAsString(i, 14);

                    Console.WriteLine(mess);
                    // Add some text to file    
                    adsFrame.Clear();
                    if (packingFrame == "C")
                    {
                        adsFrame.Add("CONCARTONE");
                    }// c concartone

                    if (motorFrame == "T2" || motorFrame == "T3" || motorFrame == "T6" || motorFrame == "T56")
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


                    Console.WriteLine(adsFrame);
                    for (int k = 0; k < adsFrame.Count; k++)
                    {

                        if (ad1.Length + adsFrame[k].Length <= 16)
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

                    
                    for (int j = 1; j <= piecesFrame; j++)
                    {
                      // frameLabelFinalFile = new SLDocument(shrani);

                        SLStyle odebeljeno = frameLabelFinalFile.CreateStyle();
                        odebeljeno.Font.Bold = true;
                        odebeljeno.Font.FontName = "Arial CE";


                        //prva vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 1, "ORDINE:");
                        frameLabelFinalFile.SetCellValue(stevec, 3, ordineFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7, stickerDate);
                        frameLabelFinalFile.SetCellValue(stevec, 1 + 13, "ORDINE:");
                        frameLabelFinalFile.SetCellValue(stevec, 3 + 13, ordineFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7 + 13, stickerDate);
                        stevec++;

                        //druga vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 1 + 13, "RIF:");
                        frameLabelFinalFile.SetCellValue(stevec, 2, deliveyCompanyFrame + rifFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 1 + 13, "RIF:");
                        frameLabelFinalFile.SetCellValue(stevec, 2 + 13, deliveyCompanyFrame + rifFrame);
                        stevec++;


                        //tretja vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 7, mountTypeFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7 + 13, mountTypeFrame);
                        if (mountTypeFrame == "CB")
                        {
                            frameLabelFinalFile.SetCellValue(stevec, 1, "CON FORI X MECCANISMO CONFORT");
                            frameLabelFinalFile.SetCellValue(stevec, 1 + 13, "CON FORI X MECCANISMO CONFORT");
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
                            frameLabelFinalFile.SetCellStyle(stevec, 1 + 13, fontMereLength);
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
                        Console.WriteLine(fontMereLength.Font.FontSize);
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
                        frameLabelFinalFile.SetCellStyle(stevec, 1 + 13, odebeljeno);
                        //desetavrstica
                        
                        //barCodeTest = ordineFrame+"-"+italCodeFrame + "-" + modelFrame + "-" + "RETE";
                        try
                        {
                            Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                            barCodeTest = ordineFrame + "-" + italCodeFrame + "-" + modelFrame + "-" + "RETE";
                            Image image = brCode.Draw(barCodeTest, 30); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali
                                                                        // image.Save("frameBarCode.gif");
                            images.Add(image);

                            images[barCounter].Save("frameBarCode.gif");
                            SLPicture pic = new SLPicture("frameBarCode.gif");
                            pic.SetPosition(stevec, 1);
                            frameLabelFinalFile.InsertPicture(pic);
                            pic.SetPosition(stevec, 14);
                            frameLabelFinalFile.InsertPicture(pic);
                            barCounter++;

                            //barCodeIndex = barCodeIndex + 9;

                            //frameLabelFinalFile.Save();
                            barCodeTest = "";
                            Console.WriteLine(modelFrame);
                            stevec++;

                            //frameLabelFinalFile.Save();

                        }
                        catch
                        {

                        }
                        
                        //to bos izbrisal drugic
                        stevec = stevec + 2;

                        Console.WriteLine("nekar me");
                    }
                    ad1 = ad2 = ad3 = "";

                }



                // format save name of file to save on user destop
                //save sticker file

                //frameLabelFinalFile.CloseWithoutSaving(); //close order file


                //  if (vrsticaCheck != "")
                /*  {
                      steviloKosov = fileNarocila.GetCellValueAsInt32(i, 18);// get number of same items

                      for (int j = steviloKosov; j > 0; j--)
                      {
                          //getting stuff from order file
                          narocilo = fileNarocila.GetCellValueAsString(i, 3);
                          oznakaModel = fileNarocila.GetCellValueAsString(i, 4);
                          model = fileNarocila.GetCellValueAsString(i, 5);
                          vrsta = fileNarocila.GetCellValueAsString(i, 6);
                          guma = fileNarocila.GetCellValueAsString(i, 7);
                          sirina = fileNarocila.GetCellValueAsString(i, 8);
                          dolzina = fileNarocila.GetCellValueAsString(i, 9);
                          string mere = sirina + " x " + dolzina;
                          dodatek1 = fileNarocila.GetCellValueAsString(i, 10); //CB,AU,GA ....
                          dodatek2 = fileNarocila.GetCellValueAsString(i, 11);  //karton
                          dodatek3 = fileNarocila.GetCellValueAsString(i, 12); // vreča
                          dodatek4 = fileNarocila.GetCellValueAsString(i, 13); // dodatek 60*40 + new model
                          kraj = fileNarocila.GetCellValueAsString(i, 14);
                          rif = fileNarocila.GetCellValueAsString(i, 15);
                          rifInKraj = kraj + rif;
                          lastnost1 = fileNarocila.GetCellValueAsString(i, 16);
                          lastnost1 = lastnost1.Replace("-1CM", "");
                          lastnost2 = fileNarocila.GetCellValueAsString(i, 17);

                          if (dodatek2 == "C" && dodatek3 != "D") // preveriz "C" za krton;
                          {
                              fileNalepke.SetCellValue(kazalec + 7, 4, "CON CARTONE");
                              fileNalepke.SetCellValue(kazalec + 7, 17, "CON CARTONE");
                          }
                          if (dodatek2 != "C" && dodatek3 == "D")
                          {
                              fileNalepke.SetCellValue(kazalec + 7, 4, "SACHETTO");
                              fileNalepke.SetCellValue(kazalec + 7, 17, "SACHETTO");
                          }
                          if (dodatek2 == "C" && dodatek3 == "D")
                          {
                              fileNalepke.SetCellValue(kazalec + 7, 4, "SACHETTO+CARTONE");
                              fileNalepke.SetCellValue(kazalec + 7, 17, "SACHETTO+CARTONE");
                          }

                          fileNalepke.SetCellValue(kazalec, 1, "ORDINE:"); //ordinare box, rif box , date box, supplement box (cb)
                          fileNalepke.SetCellValue(kazalec, 14, "ORDINE:");
                          if (rifInKraj != "")
                          {
                              fileNalepke.SetCellValue(kazalec + 1, 1, "RIF.");
                              fileNalepke.SetCellValue(kazalec + 1, 14, "RIF.");
                              fileNalepke.SetCellValue(kazalec + 1, 2, rifInKraj);
                              fileNalepke.SetCellValue(kazalec + 1, 15, rifInKraj);
                          }

                          fileNalepke.SetCellValue(kazalec + 2, 7, dodatek1);
                          fileNalepke.SetCellValue(kazalec + 2, 20, dodatek1);
                          fileNalepke.SetCellValue(kazalec, 7, formatiranDatum);
                          fileNalepke.SetCellValue(kazalec, 20, formatiranDatum);
                          if (dodatek1 == "CB")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, cbText1); //CB text add
                              fileNalepke.SetCellValue(kazalec + 3, 1, cbText2); // podaj narekovaje
                              fileNalepke.SetCellValue(kazalec + 2, 14, cbText1);
                              fileNalepke.SetCellValue(kazalec + 3, 14, cbText2); // podaj narekovaje
                          }
                          if (dodatek1 == "CC")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, ccText1); //CC text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, ccText1);
                          }
                          if (dodatek1 == "CL")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, clText1); //CL text add
                              fileNalepke.SetCellValue(kazalec + 3, 1, clText2);
                              fileNalepke.SetCellValue(kazalec + 2, 14, clText1);
                              fileNalepke.SetCellValue(kazalec + 3, 14, clText2);
                          }
                          if (dodatek1 == "LF")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, lfText1); //LF text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, lfText1);
                          }
                          if (dodatek1 == "GO")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, goText1); //GO text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, goText1);
                          }
                          if (dodatek1 == "AU")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, auText1); //AU text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, auText1);
                          }


                          fileNalepke.SetCellValue(kazalec, 3, narocilo); //narocilo, rifkraj
                          fileNalepke.SetCellValue(kazalec, 16, narocilo);


                          if (model == "EVO   SATURNO")
                          {
                              fileNalepke.SetCellValue(kazalec + 4, 1, "SATURNO"); //saturno check model and type
                              fileNalepke.SetCellValue(kazalec + 4, 7, "E1");
                              fileNalepke.SetCellValue(kazalec + 4, 14, "SATURNO");
                              fileNalepke.SetCellValue(kazalec + 4, 20, "E1");
                              fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                              fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);

                              fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                              fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                              fileNalepke.SetCellValue(kazalec + 8, 1, " ");
                              fileNalepke.SetCellValue(kazalec + 8, 14, " ");
                          }
                          if (model == "EVO  PT  PLUTONE")
                          {
                              fileNalepke.SetCellValue(kazalec + 4, 1, "PLUTONE"); // plutone check model and type
                              fileNalepke.SetCellValue(kazalec + 4, 7, "E2");
                              fileNalepke.SetCellValue(kazalec + 4, 14, "PLUTONE");
                              fileNalepke.SetCellValue(kazalec + 4, 20, "E2");
                              fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                              fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);

                              fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                              fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                              fileNalepke.SetCellValue(kazalec + 8, 1, " ");
                              fileNalepke.SetCellValue(kazalec + 8, 14, " ");
                          }
                          if (model == "EVO  PT  NETTUNO")
                          {
                              fileNalepke.SetCellValue(kazalec + 4, 1, "NETTUNO"); //nettuno check model and type
                              fileNalepke.SetCellValue(kazalec + 4, 7, "E3");
                              fileNalepke.SetCellValue(kazalec + 4, 14, "NETTUNO");
                              fileNalepke.SetCellValue(kazalec + 4, 20, "E3");
                              fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                              fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                              fileNalepke.SetCellValue(kazalec + 8, 1, " ");
                              fileNalepke.SetCellValue(kazalec + 8, 14, " ");
                              if (dodatek1 == "")
                              {
                                  fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                  fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                  fileNalepke.SetCellValue(kazalec + 8, 1, motoriNapis);
                                  fileNalepke.SetCellValue(kazalec + 8, 14, motoriNapis);
                              }
                              if (lastnost2.Contains("T1") || lastnost2.Contains("T2") || lastnost2.Contains("T3") || lastnost2.Contains("T4") || lastnost2.Contains("T5"))
                              {
                                  fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                  fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                  fileNalepke.SetCellValue(kazalec + 8, 1, motoriNapis);
                                  fileNalepke.SetCellValue(kazalec + 8, 14, motoriNapis);
                              }

                              fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                              fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);


                          }
                          if (model != "EVO   SATURNO" && model != "EVO  PT  PLUTONE" && model != "EVO  PT  NETTUNO")
                          {
                                  fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                                  fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);

                              if (lastnost2.Contains("T1") || lastnost2.Contains("T2") || lastnost2.Contains("T3") || lastnost2.Contains("T4") || lastnost2.Contains("T5"))
                              {
                                  if (lastnost2.StartsWith("NCT3"))
                                  {
                                      if (lastnost1 == "")
                                      {
                                          fileNalepke.SetCellValue(kazalec + 6, 4, lastnost2);
                                          fileNalepke.SetCellValue(kazalec + 6, 17, lastnost2);
                                      }
                                      else
                                      {
                                          MessageBox.Show("Nekje je napaka");
                                      }
                                  }
                                  else
                                  {
                                      fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                      fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                      fileNalepke.SetCellValue(kazalec + 8, 1, motoriNapis + "    " + dodatek4);
                                      fileNalepke.SetCellValue(kazalec + 8, 14, motoriNapis + "    " + dodatek4);
                                  }

                              }
                              else
                              {
                                  fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                  fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                  fileNalepke.SetCellValue(kazalec + 8, 1, lastnost2 + "    "+ dodatek4);
                                  fileNalepke.SetCellValue(kazalec + 8, 14, lastnost2 + "    " +dodatek4);
                              }

                              if (model.Length <= 9)
                              {
                                  fileNalepke.SetCellValue(kazalec + 4, 1, model); // other model's thab saturno, plutone, nettuno -> guma, type, model
                                  fileNalepke.SetCellValue(kazalec + 4, 14, model);
                              }
                              else
                              {
                                  if (model.Length > 9 && model.Length < 13)
                                  {
                                      fileNalepke.SetCellStyle(kazalec + 4, 1, fontModelLength1);
                                      fileNalepke.SetCellStyle(kazalec + 4, 14, fontModelLength1);
                                      fileNalepke.SetCellValue(kazalec + 4, 1, model); // other model's thab saturno, plutone, nettuno -> guma, type, model
                                      fileNalepke.SetCellValue(kazalec + 4, 14, model);
                                  }
                                  if (model.Length >= 13 && model.Length < 17)
                                  {
                                      fileNalepke.SetCellStyle(kazalec + 4, 1, fontModelLength2);
                                      fileNalepke.SetCellStyle(kazalec + 4, 14, fontModelLength2);
                                      fileNalepke.SetCellValue(kazalec + 4, 1, model); // other model's thab saturno, plutone, nettuno -> guma, type, model
                                      fileNalepke.SetCellValue(kazalec + 4, 14, model);
                                  }
                              }

                                  fileNalepke.SetCellValue(kazalec + 4, 7, vrsta);
                                  fileNalepke.SetCellValue(kazalec + 4, 20, vrsta);
                                  fileNalepke.SetCellValue(kazalec + 4, 6, guma);
                                  fileNalepke.SetCellValue(kazalec + 4, 19, guma);
                              }

                          if (mere.Length > 9) {
                              fileNalepke.SetCellStyle(kazalec + 6, 1, fontMereLength);
                              fileNalepke.SetCellStyle(kazalec + 6, 14, fontMereLength);
                              fileNalepke.SetCellValue(kazalec + 6, 1, mere); //sizes 
                              fileNalepke.SetCellValue(kazalec + 6, 14, mere);
                          }
                          else
                          {
                              fileNalepke.SetCellValue(kazalec + 6, 1, mere); //sizes 
                              fileNalepke.SetCellValue(kazalec + 6, 14, mere);
                          }


                              if (oznakaModel != "") {
                              fileNalepke.SetCellValue(kazalec + 4, 5, oznakaModel);
                              fileNalepke.SetCellValue(kazalec + 4, 18, oznakaModel);
                              }


                              kazalec += 11; // next sticker pointer
                          }

                  } */


                //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //get current user destop path
                //                string shrani = pathPredloga + "\\nalepkeProgram\\" + datum + " NALEPKE.xlsx";
                //string shrani = path + "\\" + datum + "NALEPKE.xlsx"; // format save name of file to save on user destop
                //MessageBox.Show(shrani);
                //                 fileNalepke.SaveAs(shrani); //save sticker file
                

                for (int haha = 0; haha <= (images.Count()-1); haha++)
                {
                    images[haha].Save("frameBarCode.gif");
                    SLPicture pic = new SLPicture("frameBarCode.gif");
                    pic.SetPosition(barCodeIndex, 1);
                    frameLabelFinalFile.InsertPicture(pic);
                    pic.SetPosition(barCodeIndex, 14);
                    frameLabelFinalFile.InsertPicture(pic);
                    barCodeIndex = barCodeIndex + 9;
                }
                MessageBox.Show(shrani);
                frameLabelFinalFile.SaveAs(shrani);
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
                            Console.WriteLine(datoteka);
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

                Console.WriteLine(lastRowBedIndex);

                for (int i = 4; i < lastRowBedIndex; i++)
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


                    Console.WriteLine(modelBed + bedOrderNumber);
                }





                Console.WriteLine("tukaj");
                Console.WriteLine(datoteka);
                //SLDocument fileNalepke = new SLDocument("template.xlsx");// open template file
                string pathPredloga = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string kocnoPredlogaPath = pathPredloga + "\\nalepkeProgram\\nalepke.xlsx";
                //SLDocument fileNalepke = new SLDocument(kocnoPredlogaPath);
                //SLDocument fileNalepkeOblecene = new SLDocument("predlogaOblecene.xlsx");
                //     SLStyle fontMereLength = fileNalepke.CreateStyle();
                //    SLStyle fontModelLength1 = fileNalepke.CreateStyle();
                //    SLStyle odebeljeno = fileNalepke.CreateStyle();

                /*    fontMereLength.Font.FontSize = 14; // change if string lengt of size is too long 9 chars.
                    fontModelLength1.Font.FontSize = 16;
                    fontMereLength.Font.FontName = "Arial CE";
                    fontModelLength1.Font.FontName = "Arial CE";
                    odebeljeno.Font.Bold = true;
                    odebeljeno.Font.FontName = "Arial CE";

                    SLStyle fontModelLength2 = fileNalepke.CreateStyle(); 
                    fontModelLength2.Font.FontSize = 13.5;
                    fontModelLength2.Font.FontName = "Arial CE"; 

                Random r = new Random(); //kaj je to ?

                    int kazalec = 2;

                    string datum = fileNarocila.GetCellValueAsString("F1"); // get date and format it 
                    string[] splitanDatum = datum.Split('.');
                    string formatiranDatum = "";
                    if (splitanDatum[0].Length != 2)
                    {
                        formatiranDatum += "0" + splitanDatum[1];
                    }
                    else
                    {
                        formatiranDatum += splitanDatum[1];
                    }
                    string leto = splitanDatum[2].Substring((splitanDatum[2].Length - 2), 2);
                    formatiranDatum += leto; */


                //MessageBox.Show(steviloStrani.ToString());




                //  if (vrsticaCheck != "")
                /*  {
                      steviloKosov = fileNarocila.GetCellValueAsInt32(i, 18);// get number of same items

                      for (int j = steviloKosov; j > 0; j--)
                      {
                          //getting stuff from order file
                          narocilo = fileNarocila.GetCellValueAsString(i, 3);
                          oznakaModel = fileNarocila.GetCellValueAsString(i, 4);
                          model = fileNarocila.GetCellValueAsString(i, 5);
                          vrsta = fileNarocila.GetCellValueAsString(i, 6);
                          guma = fileNarocila.GetCellValueAsString(i, 7);
                          sirina = fileNarocila.GetCellValueAsString(i, 8);
                          dolzina = fileNarocila.GetCellValueAsString(i, 9);
                          string mere = sirina + " x " + dolzina;
                          dodatek1 = fileNarocila.GetCellValueAsString(i, 10); //CB,AU,GA ....
                          dodatek2 = fileNarocila.GetCellValueAsString(i, 11);  //karton
                          dodatek3 = fileNarocila.GetCellValueAsString(i, 12); // vreča
                          dodatek4 = fileNarocila.GetCellValueAsString(i, 13); // dodatek 60*40 + new model
                          kraj = fileNarocila.GetCellValueAsString(i, 14);
                          rif = fileNarocila.GetCellValueAsString(i, 15);
                          rifInKraj = kraj + rif;
                          lastnost1 = fileNarocila.GetCellValueAsString(i, 16);
                          lastnost1 = lastnost1.Replace("-1CM", "");
                          lastnost2 = fileNarocila.GetCellValueAsString(i, 17);

                          if (dodatek2 == "C" && dodatek3 != "D") // preveriz "C" za krton;
                          {
                              fileNalepke.SetCellValue(kazalec + 7, 4, "CON CARTONE");
                              fileNalepke.SetCellValue(kazalec + 7, 17, "CON CARTONE");
                          }
                          if (dodatek2 != "C" && dodatek3 == "D")
                          {
                              fileNalepke.SetCellValue(kazalec + 7, 4, "SACHETTO");
                              fileNalepke.SetCellValue(kazalec + 7, 17, "SACHETTO");
                          }
                          if (dodatek2 == "C" && dodatek3 == "D")
                          {
                              fileNalepke.SetCellValue(kazalec + 7, 4, "SACHETTO+CARTONE");
                              fileNalepke.SetCellValue(kazalec + 7, 17, "SACHETTO+CARTONE");
                          }

                          fileNalepke.SetCellValue(kazalec, 1, "ORDINE:"); //ordinare box, rif box , date box, supplement box (cb)
                          fileNalepke.SetCellValue(kazalec, 14, "ORDINE:");
                          if (rifInKraj != "")
                          {
                              fileNalepke.SetCellValue(kazalec + 1, 1, "RIF.");
                              fileNalepke.SetCellValue(kazalec + 1, 14, "RIF.");
                              fileNalepke.SetCellValue(kazalec + 1, 2, rifInKraj);
                              fileNalepke.SetCellValue(kazalec + 1, 15, rifInKraj);
                          }

                          fileNalepke.SetCellValue(kazalec + 2, 7, dodatek1);
                          fileNalepke.SetCellValue(kazalec + 2, 20, dodatek1);
                          fileNalepke.SetCellValue(kazalec, 7, formatiranDatum);
                          fileNalepke.SetCellValue(kazalec, 20, formatiranDatum);
                          if (dodatek1 == "CB")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, cbText1); //CB text add
                              fileNalepke.SetCellValue(kazalec + 3, 1, cbText2); // podaj narekovaje
                              fileNalepke.SetCellValue(kazalec + 2, 14, cbText1);
                              fileNalepke.SetCellValue(kazalec + 3, 14, cbText2); // podaj narekovaje
                          }
                          if (dodatek1 == "CC")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, ccText1); //CC text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, ccText1);
                          }
                          if (dodatek1 == "CL")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, clText1); //CL text add
                              fileNalepke.SetCellValue(kazalec + 3, 1, clText2);
                              fileNalepke.SetCellValue(kazalec + 2, 14, clText1);
                              fileNalepke.SetCellValue(kazalec + 3, 14, clText2);
                          }
                          if (dodatek1 == "LF")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, lfText1); //LF text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, lfText1);
                          }
                          if (dodatek1 == "GO")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, goText1); //GO text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, goText1);
                          }
                          if (dodatek1 == "AU")
                          {
                              fileNalepke.SetCellValue(kazalec + 2, 1, auText1); //AU text add
                              fileNalepke.SetCellValue(kazalec + 2, 14, auText1);
                          }


                          fileNalepke.SetCellValue(kazalec, 3, narocilo); //narocilo, rifkraj
                          fileNalepke.SetCellValue(kazalec, 16, narocilo);


                          if (model == "EVO   SATURNO")
                          {
                              fileNalepke.SetCellValue(kazalec + 4, 1, "SATURNO"); //saturno check model and type
                              fileNalepke.SetCellValue(kazalec + 4, 7, "E1");
                              fileNalepke.SetCellValue(kazalec + 4, 14, "SATURNO");
                              fileNalepke.SetCellValue(kazalec + 4, 20, "E1");
                              fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                              fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);

                              fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                              fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                              fileNalepke.SetCellValue(kazalec + 8, 1, " ");
                              fileNalepke.SetCellValue(kazalec + 8, 14, " ");
                          }
                          if (model == "EVO  PT  PLUTONE")
                          {
                              fileNalepke.SetCellValue(kazalec + 4, 1, "PLUTONE"); // plutone check model and type
                              fileNalepke.SetCellValue(kazalec + 4, 7, "E2");
                              fileNalepke.SetCellValue(kazalec + 4, 14, "PLUTONE");
                              fileNalepke.SetCellValue(kazalec + 4, 20, "E2");
                              fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                              fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);

                              fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                              fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                              fileNalepke.SetCellValue(kazalec + 8, 1, " ");
                              fileNalepke.SetCellValue(kazalec + 8, 14, " ");
                          }
                          if (model == "EVO  PT  NETTUNO")
                          {
                              fileNalepke.SetCellValue(kazalec + 4, 1, "NETTUNO"); //nettuno check model and type
                              fileNalepke.SetCellValue(kazalec + 4, 7, "E3");
                              fileNalepke.SetCellValue(kazalec + 4, 14, "NETTUNO");
                              fileNalepke.SetCellValue(kazalec + 4, 20, "E3");
                              fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                              fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                              fileNalepke.SetCellValue(kazalec + 8, 1, " ");
                              fileNalepke.SetCellValue(kazalec + 8, 14, " ");
                              if (dodatek1 == "")
                              {
                                  fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                  fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                  fileNalepke.SetCellValue(kazalec + 8, 1, motoriNapis);
                                  fileNalepke.SetCellValue(kazalec + 8, 14, motoriNapis);
                              }
                              if (lastnost2.Contains("T1") || lastnost2.Contains("T2") || lastnost2.Contains("T3") || lastnost2.Contains("T4") || lastnost2.Contains("T5"))
                              {
                                  fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                  fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                  fileNalepke.SetCellValue(kazalec + 8, 1, motoriNapis);
                                  fileNalepke.SetCellValue(kazalec + 8, 14, motoriNapis);
                              }

                              fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                              fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);


                          }
                          if (model != "EVO   SATURNO" && model != "EVO  PT  PLUTONE" && model != "EVO  PT  NETTUNO")
                          {
                                  fileNalepke.SetCellValue(kazalec + 6, 4, lastnost1); // lastnosti 
                                  fileNalepke.SetCellValue(kazalec + 6, 17, lastnost1);

                              if (lastnost2.Contains("T1") || lastnost2.Contains("T2") || lastnost2.Contains("T3") || lastnost2.Contains("T4") || lastnost2.Contains("T5"))
                              {
                                  if (lastnost2.StartsWith("NCT3"))
                                  {
                                      if (lastnost1 == "")
                                      {
                                          fileNalepke.SetCellValue(kazalec + 6, 4, lastnost2);
                                          fileNalepke.SetCellValue(kazalec + 6, 17, lastnost2);
                                      }
                                      else
                                      {
                                          MessageBox.Show("Nekje je napaka");
                                      }
                                  }
                                  else
                                  {
                                      fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                      fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                      fileNalepke.SetCellValue(kazalec + 8, 1, motoriNapis + "    " + dodatek4);
                                      fileNalepke.SetCellValue(kazalec + 8, 14, motoriNapis + "    " + dodatek4);
                                  }

                              }
                              else
                              {
                                  fileNalepke.SetCellStyle(kazalec + 8, 1, odebeljeno);
                                  fileNalepke.SetCellStyle(kazalec + 8, 14, odebeljeno);
                                  fileNalepke.SetCellValue(kazalec + 8, 1, lastnost2 + "    "+ dodatek4);
                                  fileNalepke.SetCellValue(kazalec + 8, 14, lastnost2 + "    " +dodatek4);
                              }

                              if (model.Length <= 9)
                              {
                                  fileNalepke.SetCellValue(kazalec + 4, 1, model); // other model's thab saturno, plutone, nettuno -> guma, type, model
                                  fileNalepke.SetCellValue(kazalec + 4, 14, model);
                              }
                              else
                              {
                                  if (model.Length > 9 && model.Length < 13)
                                  {
                                      fileNalepke.SetCellStyle(kazalec + 4, 1, fontModelLength1);
                                      fileNalepke.SetCellStyle(kazalec + 4, 14, fontModelLength1);
                                      fileNalepke.SetCellValue(kazalec + 4, 1, model); // other model's thab saturno, plutone, nettuno -> guma, type, model
                                      fileNalepke.SetCellValue(kazalec + 4, 14, model);
                                  }
                                  if (model.Length >= 13 && model.Length < 17)
                                  {
                                      fileNalepke.SetCellStyle(kazalec + 4, 1, fontModelLength2);
                                      fileNalepke.SetCellStyle(kazalec + 4, 14, fontModelLength2);
                                      fileNalepke.SetCellValue(kazalec + 4, 1, model); // other model's thab saturno, plutone, nettuno -> guma, type, model
                                      fileNalepke.SetCellValue(kazalec + 4, 14, model);
                                  }
                              }

                                  fileNalepke.SetCellValue(kazalec + 4, 7, vrsta);
                                  fileNalepke.SetCellValue(kazalec + 4, 20, vrsta);
                                  fileNalepke.SetCellValue(kazalec + 4, 6, guma);
                                  fileNalepke.SetCellValue(kazalec + 4, 19, guma);
                              }

                          if (mere.Length > 9) {
                              fileNalepke.SetCellStyle(kazalec + 6, 1, fontMereLength);
                              fileNalepke.SetCellStyle(kazalec + 6, 14, fontMereLength);
                              fileNalepke.SetCellValue(kazalec + 6, 1, mere); //sizes 
                              fileNalepke.SetCellValue(kazalec + 6, 14, mere);
                          }
                          else
                          {
                              fileNalepke.SetCellValue(kazalec + 6, 1, mere); //sizes 
                              fileNalepke.SetCellValue(kazalec + 6, 14, mere);
                          }


                              if (oznakaModel != "") {
                              fileNalepke.SetCellValue(kazalec + 4, 5, oznakaModel);
                              fileNalepke.SetCellValue(kazalec + 4, 18, oznakaModel);
                              }


                              kazalec += 11; // next sticker pointer
                          }

                  } */


                //get current user destop path
                //                string shrani = pathPredloga + "\\nalepkeProgram\\" + datum + " NALEPKE.xlsx";
                //string shrani = path + "\\" + datum + "NALEPKE.xlsx"; // format save name of file to save on user destop
                //MessageBox.Show(shrani);
                //                 fileNalepke.SaveAs(shrani); //save sticker file
                // fileNarocila.CloseWithoutSaving(); //close order file
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

