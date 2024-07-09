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
using System.Globalization;
using SpreadsheetLight;

namespace nalepkeAdena
{
    public partial class Form1 : Form
    {
        //folder on desktop where app is placed
        string pathDestopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\makeSticker";

        static DateTime thisDay = DateTime.Today;
        string dateToday = thisDay.Day.ToString() + "-" + thisDay.Month.ToString();

        SLDocument document;
        SLDocument smallStickerDocument;



        string modelBed;
        string sizeXBed;
        string sizeYBed;
        string bedRowNum;
        string bedOrderNumLocal;
        string bedOrderNumber;
        string bedDeliveryCompany;
        string bedRif;
        string bedDescription;
        int piecesBed;
        string legsBedString;
        string legType;
        int legQty;
        string bedSheet;
        public int indexBed1 = 1;
        public int indexBed2 = 1;
        public int indexBed3 = 1;
        public int indexBed4 = 1;
        public int indexBed5 = 1;

        string descriptionBed1 = "";
        string descriptionBed2 = "";



        string truck;
        string datoteka = null;
        string ordineFrame;
        string barCodeTest = "";
        string deliveyCompanyFrame;
        string vVFrame;
        string rifFrame;
        string motorFrame;
        string legsFrame;
        string italCodeFrame;
        string personalizationFrame;
        string descriptionFrame;
        string descriptionFrameAddition;
        string firstAdFrame;
        string secondAdFrame;
        string thirdAdFrame;
        string forthAdFrame;
        string addsString;
        string modelFrame;
        string typeFrame; //1,2,3
        string sizeXFrame;
        string sizeYFrame;
        int piecesFrame;
        int counterTotal = 0;


        public string column1;
        public string column2;
        public string column4;
        public string column5;
        public string column3;
        string modelTotal;
        string barkoda;
        string column7;
        string column8;
        string column9;
        string column10;
        string column11;
        string column12;
        string column13;
        string column14;
        string column15;
        string column16;
        string ads17;
        string column18;
        string column19;
        string column20;
        string column21;
        string column22;
        string column23;
        string column24;
        string column25;
        string column26;
        string column27;
        string column28;
        string column29;
        string column30;
        string column31;
        string column32;
        string column33;
        string column34;
        string column35;
        string column36;
        string column37;
        string column38;
        string column39;
        string column40;
        string column41;
        string column51;

        public static DateTime todayDate = DateTime.Today;
        public string stickerDate = todayDate.ToString("MM/yy");


        public Form1()
        {
            InitializeComponent();
        }

        string removeLeadingZeros(string num)
        {
            // traverse the entire string
            for (int i = 0; i < num.Length; i++)
            {

                // check for the first non-zero character
                if (num[i] != '0')
                {
                    // return the remaining string
                    string res = num.Substring(i);
                    return res;
                }
            }

            // If the entire string is traversed
            // that means it didn't have a single
            // non-zero character, hence return "0"
            return "0";
        }
        public void smallStickerFrame(string ordine, string rif, string deliveryComp, string modelTotal, string sizeX, string sizeY, string adds1, string adds2,int i, int j)
        {
            try
            {
                smallStickerDocument.SetCellValue(i , j, ordine);
                smallStickerDocument.SetCellValue(i + 2, j, modelTotal);
                smallStickerDocument.SetCellValue(i + 5, j, removeLeadingZeros(sizeX) +"X"+ removeLeadingZeros(sizeY));
                smallStickerDocument.SetCellValue(i + 8, j, adds1+"   "+adds2);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            
        }
        public void smallStickerFrameSpecial(string ordine, string rif, string deliveryComp, string modelTotal, string sizeX, string sizeY, string adds1, string adds2, int i, int j)
        {
            try
            {
                smallStickerDocument.SetCellValue(i, j, ordine);
                smallStickerDocument.SetCellValue(i + 2, j, modelTotal);
                smallStickerDocument.SetCellValue(i + 5, j, removeLeadingZeros(sizeX) + "X" + removeLeadingZeros(sizeY));
                smallStickerDocument.SetCellValue(i + 8, j, adds1 + "   " + adds2);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

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
                Console.WriteLine("ZAČETEK");
                Console.WriteLine("IME LISTE:" + datoteka);

                string koncnaPredlogaPathFrame = pathDestopFolder + "\\templateFrame.xlsx";

                SLWorksheetStatistics stats = fileNarocila.GetWorksheetStatistics(); // stats for order file, to get last row

                SLDocument frameLabelFinalFile = new SLDocument(koncnaPredlogaPathFrame);

                string shrani = pathDestopFolder + "\\" + dateToday + "FRAME.xlsx";

                MessageBox.Show(shrani);
                frameLabelFinalFile.SaveAs(shrani); //save sticker file

                int stevec = 1;

                for (int i = 3; i <= stats.NumberOfRows; i++)
                {
                    Console.WriteLine(stats.NumberOfRows);
                    Console.WriteLine(i);
                    ordineFrame = fileNarocila.GetCellValueAsString(i, 4);
                    rifFrame = fileNarocila.GetCellValueAsString(i, 16);
                    vVFrame = fileNarocila.GetCellValueAsString(i, 8);
                    deliveyCompanyFrame = fileNarocila.GetCellValueAsString(i, 15);


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

                    personalizationFrame = fileNarocila.GetCellValueAsString(i, 5);
                    typeFrame = fileNarocila.GetCellValueAsString(i, 7);
                    sizeXFrame = fileNarocila.GetCellValueAsString(i, 9);
                    sizeYFrame = fileNarocila.GetCellValueAsString(i, 10);
                    firstAdFrame = fileNarocila.GetCellValueAsString(i, 11);
                    secondAdFrame = fileNarocila.GetCellValueAsString(i, 12);
                    thirdAdFrame = fileNarocila.GetCellValueAsString(i, 13);
                    forthAdFrame = fileNarocila.GetCellValueAsString(i, 14);
                    descriptionFrame = fileNarocila.GetCellValueAsString(i, 17);
                    descriptionFrameAddition = fileNarocila.GetCellValueAsString(i, 18);
                    legsFrame = fileNarocila.GetCellValueAsString(i, 19);
                    motorFrame = fileNarocila.GetCellValueAsString(i, 20);
                    piecesFrame = fileNarocila.GetCellValueAsInt32(i, 22);
                    italCodeFrame = fileNarocila.GetCellValueAsString(i, 1);

                    if (personalizationFrame == "EN" && modelFrame == "ESTREMA")
                    {
                        modelFrame = "EURONUIT";
                    }



                    if (modelFrame == "")
                    {
                        break;
                    }


                    if (motorFrame == "T2" || motorFrame == "T3" || motorFrame == "T6" || motorFrame == "T56")
                    {
                        addsString = addsString + " MOTORE MONTATO " + motorFrame;
                    }
                    if (firstAdFrame.Length > 2)
                    {
                        addsString = addsString + " " + firstAdFrame;
                        firstAdFrame = "";
                    }
                    if (secondAdFrame.Length > 2)
                    {
                        addsString = addsString + " " + secondAdFrame;
                        secondAdFrame = "";
                    }
                    if (thirdAdFrame.Length > 2)
                    {
                        addsString = addsString + " " + thirdAdFrame;
                        thirdAdFrame = "";
                    }
                    if (forthAdFrame.Length > 2)
                    {
                        addsString = addsString + " " + forthAdFrame;
                        forthAdFrame = "";
                    }
                    if (firstAdFrame == "C")
                    {
                        addsString = addsString + " CON CARTONE";
                        firstAdFrame = "";
                    }
                    if (secondAdFrame == "C")
                    {
                        addsString = addsString + " CON CARTONE";
                        secondAdFrame = "";
                    }
                    if (thirdAdFrame == "C")
                    {
                        addsString = addsString + " CON CARTONE";
                        thirdAdFrame = "";
                    }
                    if (forthAdFrame == "C")
                    {
                        addsString = addsString + " CON CARTONE";
                        forthAdFrame = "";
                    }


                    addsString = addsString + " " + legsFrame + " " + descriptionFrame + " " + descriptionFrameAddition;



                    for (int j = 1; j <= piecesFrame; j++)
                    {
                        frameLabelFinalFile = new SLDocument(shrani);

                        //prva vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 1, "ORDINE:");
                        frameLabelFinalFile.SetCellValue(stevec, 2, ordineFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 3, deliveyCompanyFrame + rifFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 9, stickerDate);
                        stevec++;

                        //druga vrstica

                        frameLabelFinalFile.SetCellValue(stevec, 1, modelFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 3, personalizationFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 4, vVFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 5, typeFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 6, firstAdFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 7, secondAdFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 8, thirdAdFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 9, forthAdFrame);


                        stevec++;


                        //tretja vrstica
                        frameLabelFinalFile.SetCellValue(stevec, 1, sizeXFrame + "X" + sizeYFrame);
                        frameLabelFinalFile.SetCellValue(stevec, 2, addsString);

                        try
                        {
                            Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                            barCodeTest = ordineFrame + "-" + italCodeFrame + "-" + modelFrame + "-" + "RETE";
                            //TUKAJ DEJ ZA SPACE
                            barCodeTest = System.Text.RegularExpressions.Regex.Replace(barCodeTest, @"\s+", "");

                            var image = brCode.Draw(barCodeTest, 37); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali
                            image.Save("frameBarCode.gif");
                            SLPicture pic = new SLPicture("frameBarCode.gif");
                            pic.SetPosition(stevec, 0.5);
                            frameLabelFinalFile.InsertPicture(pic);

                            barCodeTest = "";
                            Console.WriteLine(modelFrame);
                            stevec++;
                        }
                        catch
                        {

                        }
                        frameLabelFinalFile.Save();

                        //to bos izbrisal drugic
                        stevec++;

                        Console.WriteLine("nekar me");
                    }
                    addsString = "";
                }

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
        private void button5_Click(object sender, EventArgs e)//izberi datoteko
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
        private void fillCell(int counterTotal, SLDocument newList)
        {
            newList.SetCellValue(counterTotal, 1, column1);
            newList.SetCellValue(counterTotal, 2, column2);
            newList.SetCellValue(counterTotal, 3, column3);
            newList.SetCellValue(counterTotal, 4, column4);
            newList.SetCellValue(counterTotal, 5, column5);
            newList.SetCellValue(counterTotal, 6, modelTotal);
            newList.SetCellValue(counterTotal, 7, column7);
            newList.SetCellValue(counterTotal, 8, column8);
            newList.SetCellValue(counterTotal, 9, column9);
            newList.SetCellValue(counterTotal, 10, column10);
            newList.SetCellValue(counterTotal, 11, column11);
            newList.SetCellValue(counterTotal, 12, column12);
            newList.SetCellValue(counterTotal, 13, column13);
            newList.SetCellValue(counterTotal, 14, column14);
            newList.SetCellValue(counterTotal, 15, column15);
            newList.SetCellValue(counterTotal, 16, column16);
            newList.SetCellValue(counterTotal, 17, ads17);
            newList.SetCellValue(counterTotal, 18, column18);
            newList.SetCellValue(counterTotal, 19, column19);
            newList.SetCellValue(counterTotal, 20, column20);
            newList.SetCellValue(counterTotal, 21, column21);
            newList.SetCellValue(counterTotal, 22, column22);
            newList.SetCellValue(counterTotal, 23, column23);
            newList.SetCellValue(counterTotal, 24, column24);
            newList.SetCellValue(counterTotal, 25, column25);
            newList.SetCellValue(counterTotal, 26, column26);
            newList.SetCellValue(counterTotal, 27, column27);
            newList.SetCellValue(counterTotal, 28, column28);
            newList.SetCellValue(counterTotal, 29, column29);
            newList.SetCellValue(counterTotal, 30, column30);
            newList.SetCellValue(counterTotal, 31, column31);
            newList.SetCellValue(counterTotal, 32, column32);
            newList.SetCellValue(counterTotal, 33, column33);
            newList.SetCellValue(counterTotal, 34, column34);
            newList.SetCellValue(counterTotal, 35, column35);
            newList.SetCellValue(counterTotal, 36, column36);
            newList.SetCellValue(counterTotal, 37, column37);
            newList.SetCellValue(counterTotal, 38, column38);
            newList.SetCellValue(counterTotal, 39, column39);
            newList.SetCellValue(counterTotal, 40, column40);
            newList.SetCellValue(counterTotal, 41, column41);
            newList.SetCellValue(counterTotal, 51, column51);
        }
        public string GetWeekNumber()
        {
            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum.ToString();
        }
        private void fillCell1(int counterTotal, SLDocument newList, int vrstica)
        {
            newList.SetCellValue(counterTotal, 1, truck + "/" + GetWeekNumber());
            newList.SetCellValue(counterTotal, 2, column1);
            newList.SetCellValue(counterTotal, 3, vrstica);
            newList.SetCellValue(counterTotal, 4, column3);
            newList.SetCellValue(counterTotal, 5, column4);
            newList.SetCellValue(counterTotal, 6, column5);
            modelTotal = System.Text.RegularExpressions.Regex.Replace(modelTotal, @"\s+", " ");
            newList.SetCellValue(counterTotal, 7, modelTotal);
            newList.SetCellValue(counterTotal, 8, column7);
            newList.SetCellValue(counterTotal, 9, column8);
            newList.SetCellValue(counterTotal, 10, column9);
            newList.SetCellValue(counterTotal, 11, column10);
            newList.SetCellValue(counterTotal, 12, column11);
            newList.SetCellValue(counterTotal, 13, column12);
            newList.SetCellValue(counterTotal, 14, column13);
            newList.SetCellValue(counterTotal, 15, column14);
            newList.SetCellValue(counterTotal, 16, column15);
            newList.SetCellValue(counterTotal, 17, column16);
            newList.SetCellValue(counterTotal, 18, ads17);
            newList.SetCellValue(counterTotal, 20, column19);
            newList.SetCellValue(counterTotal, 21, column20);
            newList.SetCellValue(counterTotal, 22, column22);
            newList.SetCellValue(counterTotal, 23, column23);
            newList.SetCellValue(counterTotal, 24, column24);




        }
        private void potrditev_Click5(object sender, EventArgs e)
        {
            if (checkFileFormat(sender, e, datoteka))
            {

                SLDocument fileLista = new SLDocument(datoteka); //open order file
                Console.WriteLine("tukaj");
                Console.WriteLine(datoteka);
                Console.WriteLine("prvokoprvo" + fileLista.GetCellValueAsString(4, 1));



                DateTime thisDay = DateTime.Today;
                string path = "./"; //get current path
                string pathTotal = path + "\\" + "FINAL" + "NEW.xlsx";


                SLWorksheetStatistics stats = fileLista.GetWorksheetStatistics(); // stats for order file, to get last row
                SLDocument newList = new SLDocument();
                //morem glavo prekoprat
                string koncnaPredlogaPathList = pathDestopFolder + "\\templateList.xlsx";
                string shrani1 = pathDestopFolder + "\\" + dateToday + "LIST.xlsx";
                SLDocument templateXLSX = new SLDocument(koncnaPredlogaPathList);

                templateXLSX.SaveAs(shrani1);

                SLDocument totalFinalFile = new SLDocument(shrani1);


                SLStyle odebeljeno = newList.CreateStyle();
                SLStyle fontMereLength = newList.CreateStyle();
                fontMereLength.Font.FontName = "Arial CE";
                odebeljeno.Font.Bold = true;
                odebeljeno.Font.FontName = "Arial CE";

                List<string> ListBedModels = new List<string>();
                List<string> ListFrameModels = new List<string>();
                ListFrameModels.Add("SUPREMA");
                ListFrameModels.Add("ORTHOPEDIC");
                ListFrameModels.Add("ADVANCE PT");
                ListFrameModels.Add("ADVANCE");
                ListFrameModels.Add("TECNOFLEX");
                ListFrameModels.Add("MEDICAL PT");
                ListFrameModels.Add("PRIMA");
                ListFrameModels.Add("INFINITY");
                ListFrameModels.Add("NEW PERFECTA");
                ListFrameModels.Add("MEDICAL - UF");
                ListFrameModels.Add("FLEXYMED");
                ListFrameModels.Add("EVO   NETTUNO");
                ListFrameModels.Add("ESTREMA");
                ListFrameModels.Add("DYNAMIC - UF");
                ListFrameModels.Add("DYNAMIC");
                ListFrameModels.Add("SANITYMED");
                ListFrameModels.Add("TECNOFLEX - UF");
                ListFrameModels.Add("EVO   SATURNO");
                ListFrameModels.Add("EVO SATURNO PT");
                ListFrameModels.Add("EVO   PLUTONE");
                ListFrameModels.Add("EVO  PT  PLUTONE");
                ListFrameModels.Add("ADIVA");
                ListFrameModels.Add("MEDICAL PT - UF");
                ListFrameModels.Add("PROGRESS");
                ListFrameModels.Add("BASIC");
                ListFrameModels.Add("TRIOFLEX");
                ListFrameModels.Add("SANITYMED");
                ListFrameModels.Add("COMFORTFLEX");
                ListFrameModels.Add("COMFORTFLEX - UF");
                ListFrameModels.Add("ADVANCE PB");
                ListFrameModels.Add("EVOLUTION");
                ListFrameModels.Add("EVOLUTION  PT ");
                ListFrameModels.Add("EVOL.  PT ");
                ListFrameModels.Add("EVOLUTION MOSYS");
                ListFrameModels.Add("MOSYS");
                ListFrameModels.Add("ACTIVE");
                ListFrameModels.Add("ACTIVE - UF");
                ListFrameModels.Add("SPECIAL  ");
                ListFrameModels.Add("SPECIAL - UF");
                ListFrameModels.Add("SPECIAL");
                ListFrameModels.Add("ERGO - MED");
                ListFrameModels.Add("ERGO - MED - UF");
                ListFrameModels.Add("PIATTELLI");
                ListFrameModels.Add("PIATTELLI  ");
                ListFrameModels.Add("ADVANCE LARGE");
                ListFrameModels.Add("DYNAMIC LARGE");
                ListFrameModels.Add("ACTIVE LARGE");
                ListFrameModels.Add("ERGO - MED - LARGE");
                ListFrameModels.Add("ADVANCE STAND");
                ListFrameModels.Add("DYNAMIC STAND");
                ListFrameModels.Add("ACTIVE STAND");
                ListFrameModels.Add("ERGO - MED - STAND");
                ListFrameModels.Add("BONSAI PRIMA");
                ListFrameModels.Add("BONSAI ESTREMA");
                ListFrameModels.Add("BONSAI INNOVA");
                ListFrameModels.Add("BONSAI TECHNA");
                ListFrameModels.Add("BONSAI EVOLUTION ");
                ListFrameModels.Add("BONSAI - AEI");
                ListFrameModels.Add("BONSAI ERGOMED");
                ListFrameModels.Add("BONSAI SALUS");
                ListFrameModels.Add("BONSAI MEDICAL");
                ListFrameModels.Add("BONSAI EVOL  PT");
                ListFrameModels.Add("BONSAI PITTELLI");
                ListFrameModels.Add("MAGICA S.");
                ListFrameModels.Add("PRIMNEST S.");
                ListFrameModels.Add("ELEGANCE S.");
                ListFrameModels.Add("SALUS  ");
                ListFrameModels.Add("SALUS");
                ListFrameModels.Add("SALUS - UF");
                ListFrameModels.Add("INNOVA ");
                ListFrameModels.Add("INNOVA - UF");

                ListBedModels.Add("SOMMIER");
                ListBedModels.Add("HELENE");
                ListBedModels.Add("GAIA");
                ListBedModels.Add("MALIKA");
                ListBedModels.Add("ALLISON");
                ListBedModels.Add("CAMILLA");
                ListBedModels.Add("CHARLOTTA");
                ListBedModels.Add("CLAIRE");
                ListBedModels.Add("ELISABETH");
                ListBedModels.Add("GISELLE");
                ListBedModels.Add("GISELLE PLAIN");
                ListBedModels.Add("JOSEPHINE");
                ListBedModels.Add("MALIKA LARGE");
                ListBedModels.Add("JASMINE");
                ListBedModels.Add("MICHELLE");
                ListBedModels.Add("ALEXIA");
                ListBedModels.Add("SMART");
                ListBedModels.Add("DENISE");
                ListBedModels.Add("BEATRICE");
                ListBedModels.Add("NICOLE");
                ListBedModels.Add("JUSTINE ERGO");
                ListBedModels.Add("JUSTINE");
                ListBedModels.Add("CUBE");
                ListBedModels.Add("DREAM");
                ListBedModels.Add("BOOK");
                ListBedModels.Add("JUSTINE LINE ERGO");
                ListBedModels.Add("FREESTYLE");
                ListBedModels.Add("DIAMOND");
                ListBedModels.Add("PATRICYA");
                ListBedModels.Add("DIAMOND SWAR");
                ListBedModels.Add("INSIDE");
                ListBedModels.Add("ASIA");
                ListBedModels.Add("FENICE");
                ListBedModels.Add("PEGASUS");
                ListBedModels.Add("JUSTINE DOTS ERGO");
                ListBedModels.Add("JUSTINE DOTS");
                ListBedModels.Add("DORADO");
                ListBedModels.Add("ARIES ERGO");
                ListBedModels.Add("CARLOTTA");
                ListBedModels.Add("GUENDALINA");
                ListBedModels.Add("CORINNE");
                ListBedModels.Add("SIRIO");
                ListBedModels.Add("ANDROMEDA");
                ListBedModels.Add("ARIES");
                ListBedModels.Add("VEGA");
                ListBedModels.Add("IDRA");
                ListBedModels.Add("JUSTINE LINE");
                ListBedModels.Add("NIKY");
                ListBedModels.Add("MAIA");
                ListBedModels.Add("MAYA");
                ListBedModels.Add("MAYA HIGH");
                ListBedModels.Add("VICTORIA");
                // druga imena
                ListBedModels.Add("ETHIENNE");
                ListBedModels.Add("ARMONY");
                ListBedModels.Add("CHOPIN");
                ListBedModels.Add("GEORGE");
                ListBedModels.Add("MATISSE");
                ListBedModels.Add("NAUSICA");
                ListBedModels.Add("BOLT");
                ListBedModels.Add("ADELCHI");
                ListBedModels.Add("ARABESQUE");
                ListBedModels.Add("ARABESQUE PLAIN");
                ListBedModels.Add("EDWARD");
                ListBedModels.Add("JOE");
                ListBedModels.Add("OLIVER");
                ListBedModels.Add("NOA");
                ListBedModels.Add("NOA TRAPUNTATO");
                ListBedModels.Add("NOA LINE");
                ListBedModels.Add("NOA ERGO");
                ListBedModels.Add("NOA TRAPUNTATO ERGO");
                ListBedModels.Add("NOA LINE ERGO");
                ListBedModels.Add("DOROTY");
                ListBedModels.Add("DOROTY MAXI");
                ListBedModels.Add("DOLCE VITA");
                ListBedModels.Add("PORTOFINO");
                ListBedModels.Add("CLYDE");
                ListBedModels.Add("WILLIAM");
                ListBedModels.Add("VIKY");

                //druga imena novi modeli
                ListBedModels.Add("LUDWIG");
                ListBedModels.Add("GABRIEL");
                ListBedModels.Add("CONRAD");
                ListBedModels.Add("RICHARD");
                ListBedModels.Add("AXEL");
                ListBedModels.Add("THEODOR");
                ListBedModels.Add("ZEN");
                ListBedModels.Add("OSCAR");
                ListBedModels.Add("SEBASTIAN");
                ListBedModels.Add("PIERRE");
                ListBedModels.Add("JEROME");


                //new models
                ListBedModels.Add("ISABELLE");
                ListBedModels.Add("EMMA");
                ListBedModels.Add("KATE");
                ListBedModels.Add("INES");
                ListBedModels.Add("NORA");
                ListBedModels.Add("GRETA");
                ListBedModels.Add("LAILA");
                ListBedModels.Add("MIA");
                ListBedModels.Add("JULIA");
                ListBedModels.Add("EMILY");
                ListBedModels.Add("JAQUELINE");
                ListBedModels.Add("SKINNY");





                Console.WriteLine(stats.NumberOfRows);
                counterTotal = 3;
                //stajlo
                newList.SetRowStyle(1, fileLista.GetColumnStyle(1));
                newList.SetRowStyle(2, fileLista.GetColumnStyle(2));
                newList.SetRowStyle(3, fileLista.GetColumnStyle(3));
                for (int l = 1; l <= 52; l++)
                {
                    newList.SetColumnStyle(l, fileLista.GetColumnStyle(l));
                    newList.SetColumnWidth(l, fileLista.GetColumnWidth(l));
                }
                for (int i = 0; i <= stats.NumberOfRows; i++)
                {

                    Console.WriteLine(stats.NumberOfRows);
                    Console.WriteLine(i);
                    column3 = fileLista.GetCellValueAsString(i, 3);
                    column9 = fileLista.GetCellValueAsString(i, 9);
                    column10 = fileLista.GetCellValueAsString(i, 10);
                    column7 = fileLista.GetCellValueAsString(i, 7);
                    column8 = fileLista.GetCellValueAsString(i, 8);
                    DateTime today = DateTime.Today;
                    string[] collection = today.ToString("d").Split('.');

                    string modelTotalNew = "";
                    modelTotal = fileLista.GetCellValueAsString(i, 6);

                    column1 = fileLista.GetCellValueAsString(i, 1);
                    column2 = fileLista.GetCellValueAsString(i, 2);
                    column4 = fileLista.GetCellValueAsString(i, 4);
                    column5 = fileLista.GetCellValueAsString(i, 5);
                    column11 = fileLista.GetCellValueAsString(i, 11);
                    column12 = fileLista.GetCellValueAsString(i, 12);
                    column13 = fileLista.GetCellValueAsString(i, 13);
                    column14 = fileLista.GetCellValueAsString(i, 14);
                    column15 = fileLista.GetCellValueAsString(i, 15);
                    column16 = fileLista.GetCellValueAsString(i, 16);
                    ads17 = fileLista.GetCellValueAsString(i, 17);
                    column18 = fileLista.GetCellValueAsString(i, 18);
                    column19 = fileLista.GetCellValueAsString(i, 19);
                    column20 = fileLista.GetCellValueAsString(i, 20);
                    column21 = fileLista.GetCellValueAsString(i, 21).Split(',')[0];
                    column22 = fileLista.GetCellValueAsString(i, 22);
                    column23 = fileLista.GetCellValueAsString(i, 23);
                    column24 = fileLista.GetCellValueAsString(i, 24);
                    column25 = fileLista.GetCellValueAsString(i, 25);
                    column26 = fileLista.GetCellValueAsString(i, 26);
                    column27 = fileLista.GetCellValueAsString(i, 27);
                    column28 = fileLista.GetCellValueAsString(i, 28);
                    column29 = fileLista.GetCellValueAsString(i, 29);
                    column30 = fileLista.GetCellValueAsString(i, 30);
                    column31 = fileLista.GetCellValueAsString(i, 31);
                    column32 = fileLista.GetCellValueAsString(i, 32);
                    column33 = fileLista.GetCellValueAsString(i, 33);
                    column34 = fileLista.GetCellValueAsString(i, 34);
                    column35 = fileLista.GetCellValueAsString(i, 35);
                    column36 = fileLista.GetCellValueAsString(i, 36);
                    column37 = fileLista.GetCellValueAsString(i, 37);
                    column38 = fileLista.GetCellValueAsString(i, 38);
                    column39 = fileLista.GetCellValueAsString(i, 39);
                    column40 = fileLista.GetCellValueAsString(i, 40);
                    column41 = fileLista.GetCellValueAsString(i, 41);
                    column51 = fileLista.GetCellValueAsString(i, 51);


                    if (i.Equals(1) || i.Equals(2) || i.Equals(0))
                    {
                        fillCell(i, newList);
                    }
                    else
                    {
                        modelTotal.Split(' ').ToList().ForEach(Console.WriteLine);
                        if (ListBedModels.Contains(modelTotal.Split(' ')[0]))
                        {

                            if (!modelTotal.Contains("NICOLE"))
                            {
                                //SIMPLY BED
                                if (!modelTotal.Any(c => char.IsDigit(c)))
                                {
                                    //string temp = column22;
                                    //column22 = "";
                                    barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "TESTATA";
                                    barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");

                                    fillCell(counterTotal, newList);
                                    fillCell1(counterTotal, totalFinalFile, i);
                                    newList.SetCellValue(counterTotal, 52, barkoda);
                                    totalFinalFile.SetCellValue(counterTotal, 19, "TESTATA");
                                    counterTotal++;

                                    if (column21.Contains("P") || modelTotal.Contains("PIED"))
                                    {
                                        column22 = "";
                                        column1 = column21.Split('-')[2];
                                        barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "PIEDI";
                                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                        string zacasni = modelTotal;

                                        fillCell(counterTotal, newList);
                                        newList.SetCellValue(counterTotal, 52, barkoda);
                                        string numberPiedi = column21.Split('-')[1];
                                        //column23 = numberPiedi;
                                        Console.WriteLine(column22 + "SDFASDFSADFSADFSADFAS   ");
                                        modelTotal = column21.Split('-')[0];
                                        fillCell1(counterTotal, totalFinalFile, i);
                                        totalFinalFile.SetCellValue(counterTotal, 19, "PIEDI");
                                        modelTotal = zacasni;
                                        counterTotal++;
                                    }
                                    //break;
                                }
                                else
                                {
                                    //ONLY HEAD
                                    if (column10 == "")
                                    {
                                        string temp = column22;
                                        //column22 = "";
                                        barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "TESTATA";
                                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");

                                        fillCell(counterTotal, newList);
                                        fillCell1(counterTotal, totalFinalFile, i);
                                        newList.SetCellValue(counterTotal, 52, barkoda);
                                        totalFinalFile.SetCellValue(counterTotal, 19, "TESTATA");
                                        counterTotal++;

                                        if (column21.Contains("P") || modelTotal.Contains("PIED"))
                                        {
                                            column22 = "";
                                            column1 = column21.Split('-')[2];
                                            barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "PIEDI";
                                            barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                            string zacasni = modelTotal;

                                            fillCell(counterTotal, newList);
                                            newList.SetCellValue(counterTotal, 52, barkoda);
                                            string numberPiedi = column21.Split('-')[1];
                                            //column23 = numberPiedi;
                                            Console.WriteLine(column22 + "SDFASDFSADFSADFSADFAS   ");
                                            modelTotal = column21.Split('-')[0];
                                            fillCell1(counterTotal, totalFinalFile, i);
                                            totalFinalFile.SetCellValue(counterTotal, 19, "PIEDI");
                                            modelTotal = zacasni;
                                            counterTotal++;
                                        }
                                    }
                                    else
                                    {
                                        if (!modelTotal.Contains("NCT"))
                                        {
                                            if (modelTotal.Contains("3"))
                                            {
                                                string temp = column22;
                                                column22 = "";
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FODERA BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");

                                                fillCell(counterTotal, newList);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FODERA BASE");
                                                counterTotal++;
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FUSTO BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " "); fillCell(counterTotal, newList);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FUSTO BASE");
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                column22 = temp;
                                                counterTotal++;

                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "TESTATA";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " "); fillCell(counterTotal, newList);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "TESTATA");
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                counterTotal++;
                                            }
                                            if (modelTotal.Contains("4"))
                                            {
                                                string temp = column22;
                                                column22 = "";
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FODERA BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                fillCell(counterTotal, newList);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FODERA BASE");
                                                counterTotal++;
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FUSTO BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                fillCell(counterTotal, newList);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FUSTO BASE");
                                                column22 = temp;
                                                counterTotal++;

                                            }
                                            /*
                                            if (modelTotal.Contains("4") && ads17.Contains("FP"))
                                            {
                                                string temp = column22;
                                                column22 = "";
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FUSTO BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                fillCell(counterTotal, newList);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FUSTO BASE");
                                                column22 = temp;
                                                counterTotal++;
                                            }
                                            */
                                            /*
                                            if (modelTotal.Contains("3") && ads17.Contains("FP"))
                                            {
                                                string temp = column22;
                                                column22 = "";
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FUSTO BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                fillCell(counterTotal, newList);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FUSTO BASE");
                                                column22 = temp;
                                                counterTotal++;

                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "TESTATA";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                fillCell(counterTotal, newList);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "TESTATA");

                                                counterTotal++;

                                            }
                                            */
                                            if (column21.Contains("P") || modelTotal.Contains("PIED"))
                                            {
                                                column22 = "";
                                                column1 = column21.Split('-')[2];
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "PIEDI";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                string zacasni = modelTotal;

                                                fillCell(counterTotal, newList);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                string numberPiedi = column21.Split('-')[1];
                                                //column23 = numberPiedi;
                                                Console.WriteLine(column22 + "SDFASDFSADFSADFSADFAS   ");
                                                modelTotal = column21.Split('-')[0];
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "PIEDI");
                                                modelTotal = zacasni;
                                                counterTotal++;
                                            }
                                        }
                                        else
                                        {
                                            if (modelTotal.Contains("NCT4"))
                                            {
                                                string temp = column22;
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FODERA BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");

                                                column22 = "";

                                                fillCell(counterTotal, newList);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FODERA BASE");
                                                counterTotal++;
                                            }
                                            else
                                            {
                                                string temp = column22;
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FODERA BASE";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                column22 = "";
                                                fillCell(counterTotal, newList);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "FODERA BASE");
                                                counterTotal++;

                                                column22 = temp;

                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "TESTATA";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " "); fillCell(counterTotal, newList);
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "TESTATA");
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                counterTotal++;
                                            }


                                            if (column21.Contains("P") || modelTotal.Contains("PIED"))
                                            {
                                                column22 = "";
                                                column1 = column21.Split('-')[2];
                                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "PIEDI";
                                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                                string zacasni = modelTotal;

                                                fillCell(counterTotal, newList);
                                                newList.SetCellValue(counterTotal, 52, barkoda);
                                                string numberPiedi = column21.Split('-')[1];
                                                //column23 = numberPiedi;
                                                Console.WriteLine(column22 + "SDFASDFSADFSADFSADFAS   ");
                                                modelTotal = column21.Split('-')[0];
                                                fillCell1(counterTotal, totalFinalFile, i);
                                                totalFinalFile.SetCellValue(counterTotal, 19, "PIEDI");
                                                modelTotal = zacasni;
                                                counterTotal++;
                                            }
                                        }
                                    }

                                }

                            }
                            else
                            {
                                string temp = column22;
                                column22 = "";
                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FODERA BASE";
                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                fillCell(counterTotal, newList);
                                newList.SetCellValue(counterTotal, 52, barkoda);
                                fillCell1(counterTotal, totalFinalFile, i);
                                totalFinalFile.SetCellValue(counterTotal, 19, "FODERA BASE");
                                counterTotal++;

                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FUSTO BASE";
                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                fillCell(counterTotal, newList);
                                newList.SetCellValue(counterTotal, 52, barkoda);
                                fillCell1(counterTotal, totalFinalFile, i);
                                totalFinalFile.SetCellValue(counterTotal, 19, "FUSTO BASE");
                                column22 = temp;
                                counterTotal++;

                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "TESTATA";
                                barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                fillCell(counterTotal, newList);
                                newList.SetCellValue(counterTotal, 52, barkoda);
                                fillCell1(counterTotal, totalFinalFile, i);
                                totalFinalFile.SetCellValue(counterTotal, 19, "TESTATA");
                                counterTotal++;
                                if (column21.Contains("P"))
                                {
                                    column22 = "";
                                    column1 = column21.Split('-')[2];
                                    barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "PIEDI";
                                    barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                    string zacasni = modelTotal;
                                    modelTotal = column21.Split('-')[0];
                                    column22 = column21.Split('-')[1];
                                    fillCell(counterTotal, newList);
                                    newList.SetCellValue(counterTotal, 52, barkoda);

                                    fillCell1(counterTotal, totalFinalFile, i);
                                    totalFinalFile.SetCellValue(counterTotal, 19, "PIEDI");
                                    modelTotal = zacasni;
                                    counterTotal++;
                                }
                            }
                            //TREBA ŠE DOKONČAT

                        }
                        else
                        {
                            if (modelTotal.Contains("Y&M"))
                            {
                                barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "Y&M";
                                fillCell(counterTotal, newList);
                                fillCell1(counterTotal, totalFinalFile, i);
                                newList.SetCellValue(counterTotal, 52, barkoda);
                                totalFinalFile.SetCellValue(counterTotal, 19, "Y&M");
                                counterTotal++;
                            }
                            else
                            {
                                if (modelTotal.Contains("FABRIC") || modelTotal.Contains("PILLOW"))
                                {
                                    barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "FODERA BASE";
                                    fillCell(counterTotal, newList);
                                    fillCell1(counterTotal, totalFinalFile, i);
                                    newList.SetCellValue(counterTotal, 52, barkoda);
                                    totalFinalFile.SetCellValue(counterTotal, 19, "FODERA BASE");
                                    counterTotal++;
                                }
                                else
                                {
                                    barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "RETE";
                                    fillCell(counterTotal, newList);
                                    fillCell1(counterTotal, totalFinalFile, i);
                                    newList.SetCellValue(counterTotal, 52, barkoda);
                                    totalFinalFile.SetCellValue(counterTotal, 19, "RETE");
                                    counterTotal++;

                                    if (column21.Contains("P") || modelTotal.Contains("PIED"))
                                    {
                                        column22 = "";
                                        column1 = column21.Split('-')[2];
                                        barkoda = column4 + "-" + column1 + "-" + modelTotal + "-" + "PIEDI";
                                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                                        string zacasni = modelTotal;

                                        fillCell(counterTotal, newList);
                                        newList.SetCellValue(counterTotal, 52, barkoda);
                                        string numberPiedi = column21.Split('-')[1];
                                        //column23 = numberPiedi;
                                        Console.WriteLine(column22 + "SDFASDFSADFSADFSADFAS   ");
                                        modelTotal = column21.Split('-')[0];
                                        fillCell1(counterTotal, totalFinalFile, i);
                                        totalFinalFile.SetCellValue(counterTotal, 19, "PIEDI");
                                        modelTotal = zacasni;
                                        counterTotal++;
                                    }
                                }
                            }
                        }
                    }
                    Console.WriteLine("nekar me");
                }

                MessageBox.Show(pathTotal);
                newList.SaveAs(pathTotal);
                totalFinalFile.SaveAs(shrani1);
                //    fileLista.CloseWithoutSaving(); //close order file
                MessageBox.Show("Nalepke so kreirane."); //messsage shot for successful sticker create

            }

        }

        private void btnCreateLabelBed_Click(object sender, EventArgs e)
        {
            if (checkFileFormat(sender, e, datoteka))
            {
                SLDocument fileBedList = new SLDocument(datoteka); //open order file
                SLWorksheetStatistics stats = fileBedList.GetWorksheetStatistics(); // stats for order file, to get last row
                int lastRowBedIndex = stats.NumberOfRows;

                SLDocument newList = new SLDocument();

                string koncnaPredlogaPathBed = pathDestopFolder + "\\templateBed.xlsx";
                string shrani1 = pathDestopFolder + "\\" + dateToday + "BED.xlsx";
                //string shrani1 = pathDestopFolder + "\\" + dateToday + "BED.xlsx";

                SLDocument templateXLSX = new SLDocument(koncnaPredlogaPathBed);

                templateXLSX.SaveAs(shrani1);

                SLStyle odebeljeno = newList.CreateStyle();
                SLStyle fontMereLength = newList.CreateStyle();
                fontMereLength.Font.FontName = "Arial CE";
                odebeljeno.Font.Bold = true;
                odebeljeno.Font.FontName = "Arial CE";
                Console.WriteLine(lastRowBedIndex);



                for (int i = 4; i <= stats.NumberOfRows; i++)
                {
                    Console.WriteLine(stats.NumberOfRows);
                    Console.WriteLine(i);


                    bedRowNum = fileBedList.GetCellValueAsString(i, 1);
                    bedOrderNumLocal = fileBedList.GetCellValueAsString(i, 2);
                    bedOrderNumber = fileBedList.GetCellValueAsString(i, 3);
                    modelBed = fileBedList.GetCellValueAsString(i, 5);
                    sizeXBed = fileBedList.GetCellValueAsString(i, 8);
                    sizeYBed = fileBedList.GetCellValueAsString(i, 9);
                    bedDeliveryCompany = fileBedList.GetCellValueAsString(i, 14);
                    bedRif = fileBedList.GetCellValueAsString(i, 15);
                    bedDescription = fileBedList.GetCellValueAsString(i, 16);
                    piecesBed = fileBedList.GetCellValueAsInt32(i, 18);
                    legsBedString = fileBedList.GetCellValueAsString(i, 19);
                    Console.WriteLine(modelBed + bedOrderNumber);

                    string[] modelSplit = System.Text.RegularExpressions.Regex.Replace(modelBed.Trim(), @"\s+", " ").Split(' ');

                    string model;
                    string type;


                    if (modelSplit.Length == 1)
                    {
                        model = modelSplit[0];
                        type = "";
                    }
                    else if (modelBed.Contains("9-Y&M")){
                        model = modelSplit[0];
                        type = modelSplit[1];
                    }
                    else
                    {
                        model = modelSplit[0];
                        type = modelSplit[modelSplit.Length - 1];
                    }


                    // SHEET1= "FODERA BASE" SHEET2= "FUSTO BASE" SHEET3="TESTATA"

                    DateTime today = DateTime.Today;
                    string[] collection = today.ToString("d").Split('.');
                    if (bedDescription != "")
                    {
                        if (legsBedString != "")
                        {
                            legType = legsBedString.Split('-')[0];
                            legQty = Int32.Parse(legsBedString.Split('-')[1]);
                            bedSheet = "4";
                            for (int d = 0; d < legQty; d++)
                            {
                                makeStickerLegs(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, legType, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                            }
                            bedSheet = "1";
                        }

                        if (type == "" || (type.Contains("3") && !modelBed.Contains("FABRIC") && sizeYBed==""))
                        {
                            bedSheet = "3";
                            makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                            bedSheet = "1";

                        }
                        else if (modelBed.Contains("FABRIC"))
                        {
                            bedSheet = "1";
                            makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                            bedSheet = "1";
                        }
                        else
                        {

                            //if (type.Contains("3") && !bedDescription.Split(',')[0].Contains("FP") && !modelBed.Contains("FABRIC"))
                            if (type.Contains("3") && !modelBed.Contains("FABRIC"))
                            {
                                if (type.Contains("T"))
                                {
                                    bedSheet = "1";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "3";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "1";

                                }
                                else
                                {
                                    bedSheet = "1";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "2";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "5";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "3";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "1";
                                }

                            }
                            // 9.3 tukaj si dodal NICOLE! 
                            //if (type.Contains("4") && !bedDescription.Split(',')[0].Contains("FP") && !model.Contains("NICOLE"))
                            if (type.Contains("4") && !model.Contains("NICOLE"))
                            {
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "2";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "5";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "1";
                            }
                            //if (type.Contains("3") && bedDescription.Split(',')[0].Contains("FP"))
                           /* if (type.Contains("3"))
                            {
                                bedSheet = "2";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "5";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "3";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "1";
                            }
                           */
                            //if (type.Contains("4") && bedDescription.Split(',')[0].Contains("FP"))
                            if (type.Contains("4"))
                            {
                                bedSheet = "2";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "5";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "1";
                            }
                            if (model == "NICOLE")
                            {
                                if (type.Contains('4'))
                                {
                                    bedSheet = "1";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "2";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "5";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "3";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "1";
                                }
                                else
                                {
                                    bedSheet = "3";
                                    makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                    bedSheet = "1";

                                }
                            }
                            if(model == "9-Y&M" && (type=="C1" || (type == "C2") || (type == "C3") || (type == "C4") || (type == "C5") || (type == "C6"))){
                                bedSheet = "1";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "1";
                            }
                            if (model == "9-Y&M" && type == "MOD.")
                            {
                                bedSheet = "3";
                                makeSticker(bedSheet, bedRowNum, bedOrderNumber, bedDeliveryCompany, bedRif, modelBed, sizeXBed + "X" + sizeYBed, bedDescription, thisDay, piecesBed, bedOrderNumLocal);
                                bedSheet = "1";
                            }


                        }
                    }
                    Console.WriteLine("nekar me");
                }
                fileBedList.CloseWithoutSaving(); //close order file
                MessageBox.Show("Nalepke so kreirane."); //messsage shot for successful sticker create


            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            truck = comboBox1.SelectedItem.ToString();
            Console.WriteLine(truck);
            Console.WriteLine(GetWeekNumber());
        }
        public string checkType(string type, string otherAdds)
        {
            switch (type)
            {
                case "COA4+":
                    descriptionBed1 = "ALTO DRITTO 4 LATI, ERGOCOMFORT";
                    break;
                case "CO_VB3":
                    descriptionBed1 = "ALTO BOMBATO 3 LATI, ERGOCOMFORT";
                    break;
                case "NCA3+":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, NON CONT.";
                    break;
                case "NCA4+":
                    descriptionBed1 = "ALTO DRITTO 4 LATI, NON CONT.";
                    break;
                case "NCB3+":
                    descriptionBed1 = "ALTO BOMBATO 3 LATI, NON CONT.";
                    break;
                case "COA3+":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, ERGOCOMFORT";
                    break;
                case "SPB3+":
                    descriptionBed1 = "ALTO BOMBATO 3 LATI, SPACE";
                    break;
                case "SPB4+":
                    descriptionBed1 = "ALTO BOMBATO 4 LATI, SPACE";
                    break;
                case "CO_VA3+":
                    descriptionBed1 = "ALTO DRITO 3 LATI, ERGOCOMFORT";
                    break;
                case "NCM3":
                    descriptionBed1 = "ALTO DRITTO XL 3 LATI, NON CONT.";
                    break;
                case "COA4":
                    descriptionBed1 = "ALTO DRITTO 4 LATI, ERGOCOMFORT";
                    break;
                case "CO_VA3":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, ERGOCOMFORT";
                    break;
                case "SPA4":
                    descriptionBed1 = "ALTO DRITTO 4 LATI, SPACE";
                    break;
                case "SPA4+":
                    descriptionBed1 = "ALTO DRITTO 4 LATI, SPACE";
                    break;
                case "COM3":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, ERGOCOMFORT";
                    break;
                case "SPM3":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, SPACE";
                    break;
                case "COB3":
                    descriptionBed1 = "ALTO SAGOMATO 3 LATI, ERGOCOMFORT";
                    break;
                case "SPA3+":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, SPACE";
                    break;
                case "SPB3":
                    descriptionBed1 = "ALTO SAGOMATO 3 LATI, SPACE";
                    break;
                case "COA3":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, ERGOCOMFORT";
                    break;
                case "NCA3":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, NON CONT.";
                    break;
                case "SPA3":
                    descriptionBed1 = "ALTO DRITTO 3 LATI, SPACE";
                    break;
                case "NCT3":
                    descriptionBed1 = "TRAPUNTATO BASSO 3 LATI";
                    break;
                case "NCF3":
                    descriptionBed1 = "TESSILE BASSO 3 LATI";
                    break;
                case "NCS3":
                    descriptionBed1 = "TESSILE FASCIA SLIM 3 LATI";
                    break;
                case "":
                    descriptionBed1 = descriptionBed1 + "";
                    break;

            }
            return (descriptionBed1 + " " + descriptionBed2);
        }
        public void makeSticker(string sheet, string first, string bedOrderNumber, string deliveyCompany, string bedrif, string modelbed, string dimension, string adds, DateTime date, int pieces, string bedOrderNumLocal)
        {
            string shrani1 = pathDestopFolder + "\\" + dateToday + "BED.xlsx";

            for (int u = 0; u < pieces; u++)
            {
                document = new SLDocument(shrani1);
                if (sheet == "1")
                {
                    document.SelectWorksheet(sheet);

                    string modelTypeSize;
                    string[] addsSplit;
                    string fabric;
                    string[] fabricSplit;
                    string fabricType;
                    string color;

                    addsSplit = adds.Split(',');
                    fabricType = addsSplit[0].Split(' ')[0];

                    color = addsSplit[0].Split(' ')[1];




                    Console.WriteLine(fabricType + " " + color);

                    string otherAdds = "";

                    for (int k = 1; k < addsSplit.Length; k++)
                    {
                        otherAdds = addsSplit[k] + " ";
                    }
                    Console.WriteLine(otherAdds);

                    string[] modelSplit = System.Text.RegularExpressions.Regex.Replace(modelbed, @"\s+", " ").Split(' ');

                    string model = "";
                    string type;




                    if (modelSplit.Length == 1)
                    {
                        model = modelSplit[0];
                        type = "";
                        descriptionBed1 = "SIMPLY BED";

                    }
                    else
                    {
                        for (int z = 0; z < modelSplit.Length - 1; z++)
                        {
                            model += modelSplit[z] + " ";
                        }

                        type = modelSplit[modelSplit.Length - 1];
                    }


                    DateTime today = DateTime.Today;


                    checkType(type, otherAdds);

                    // XSSFFont font = wb.create;




                    //frameLabelFinalFile = new SLDocument("template.xlsx");
                    //prva vrstica
                    document.SetCellValue(indexBed1, 1, "ORDINE:");
                    document.SetCellValue(indexBed1, 3, bedOrderNumber);
                    document.SetCellValue(indexBed1, 7, bedDeliveryCompany + bedRif);
                    document.SetCellValue(indexBed1, 10, stickerDate);
                    document.SetCellValue(indexBed1, 13, bedOrderNumLocal);
                    indexBed1++;


                    modelTypeSize = model + " " + type + "    " + dimension;
                    if (modelTypeSize.Length > 28)
                    {
                        SLStyle smallerFontNameSize = document.CreateStyle();
                        smallerFontNameSize.Font.FontSize = 11;
                        document.SetCellStyle(indexBed1, 1, smallerFontNameSize);
                    }
                    document.SetCellValue(indexBed1, 1, modelTypeSize);

                    document.SetCellValue(indexBed1, 8, descriptionBed1);
                    indexBed1++;



                    document.SetCellValue(indexBed1, 1, adds);
                    document.SetCellValue(indexBed1 + 1, 1, bedOrderNumLocal.Split('/')[bedOrderNumLocal.Split('/').Length-1]);

                    try
                    {

                        Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                        barkoda = bedOrderNumber + "-" + first + "-" + modelbed + "-" + "FODERA BASE";
                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", "");
                        Image image = brCode.Draw(barkoda, 37); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali
                        image.Save("BedBarCode.gif");                                            // image.Save("frameBarCode.gif");

                        SLPicture pic = new SLPicture("BedBarCode.gif");
                        pic.SetPosition(indexBed1, 0.5);
                        document.InsertPicture(pic);

                        barkoda = "";

                    }
                    catch
                    {

                    }
                    document.Save();
                    indexBed1++;
                    indexBed1++;
                }
                if (sheet == "2")
                {
                    document.SelectWorksheet(sheet);


                    string[] addsSplit;
                    string fabricType;
                    string modelTypeSize;

                    addsSplit = adds.Split(',');
                    fabricType = addsSplit[0].Split(' ')[0];




                    string otherAdds = "";

                    for (int k = 1; k < addsSplit.Length; k++)
                    {
                        otherAdds = addsSplit[k] + " ";
                    }
                    Console.WriteLine(otherAdds);
                    string[] modelSplit = System.Text.RegularExpressions.Regex.Replace(modelbed, @"\s+", " ").Split(' ');

                    string model = "";
                    string type;


                    if (modelSplit.Length == 1)
                    {
                        model = modelSplit[0];
                        type = "";
                        descriptionBed1 = "SIMPLY BED";

                    }
                    else
                    {
                        for (int z = 0; z < modelSplit.Length - 1; z++)
                        {
                            model += modelSplit[z] + " ";
                        }

                        type = modelSplit[modelSplit.Length - 1];
                    }




                    checkType(type, otherAdds);

                    // XSSFFont font = wb.create;




                    //frameLabelFinalFile = new SLDocument("template.xlsx");
                    //prva vrstica
                    document.SetCellValue(indexBed2, 1, "ORDINE:");
                    document.SetCellValue(indexBed2, 3, bedOrderNumber);
                    document.SetCellValue(indexBed2, 7, bedDeliveryCompany + bedRif);
                    document.SetCellValue(indexBed2, 10, stickerDate);
                    document.SetCellValue(indexBed2, 13, bedOrderNumLocal);
                    indexBed2++;


                    modelTypeSize = model + " " + type + "    " + dimension;
                    if (modelTypeSize.Length > 28)
                    {
                        SLStyle smallerFontNameSize = document.CreateStyle();
                        smallerFontNameSize.Font.FontSize = 11;
                        document.SetCellStyle(indexBed2, 1, smallerFontNameSize);
                    }
                    document.SetCellValue(indexBed2, 1, modelTypeSize);

                    document.SetCellValue(indexBed2, 8, descriptionBed1);
                    indexBed2++;


                    //POSEBNE OZNAKE ZA KLANFANJE, DA ZNAJO KAJ LANFAT
                    if ((modelTypeSize.Contains("DREAM") || modelTypeSize.Contains("INSIDE") || modelTypeSize.Contains("DIAMOND") || modelTypeSize.Contains("CUBE")) && adds.Contains("FP"))
                    {
                        //adds = adds +"   000000";
                        adds = adds + "   " + (char)127 + (char)127 + (char)127 + (char)127 + (char)127 + (char)127;

                    }

                    document.SetCellValue(indexBed2, 1, adds);


                    try
                    {
                        Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                        barkoda = bedOrderNumber + "-" + first + "-" + modelbed + "-" + "FUSTO BASE";
                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", "");
                        Image image = brCode.Draw(barkoda, 37); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali
                        image.Save("BedBarCode.gif");                                            // image.Save("frameBarCode.gif");

                        SLPicture pic = new SLPicture("BedBarCode.gif");
                        pic.SetPosition(indexBed2, 1);
                        document.InsertPicture(pic);
                        //pic.SetPosition(odkod2, 14);
                        //document.InsertPicture(pic);


                        //barCodeIndex = barCodeIndex + 9;


                        barkoda = "";




                    }
                    catch
                    {

                    }
                    document.Save();
                    indexBed2++;
                    indexBed2++;
                }
                if (sheet == "3")
                {
                    document.SelectWorksheet(sheet);


                    string[] addsSplit;
                    string fabricType;
                    string color;
                    string modelTypeSize;

                    addsSplit = adds.Split(',');
                    fabricType = addsSplit[0].Split(' ')[0];
                    color = addsSplit[0].Split(' ')[1];




                    Console.WriteLine(fabricType + " " + color);

                    string otherAdds = "";

                    for (int k = 1; k < addsSplit.Length; k++)
                    {
                        otherAdds = addsSplit[k] + " ";
                    }
                    Console.WriteLine(otherAdds);

                    string[] modelSplit = System.Text.RegularExpressions.Regex.Replace(modelbed, @"\s+", " ").Split(' ');

                    string model = "";
                    string type;
                    descriptionBed1 = "";


                    if (modelSplit.Length == 1)
                    {
                        model = modelSplit[0];
                        descriptionBed1 = "SIMPLY BED";
                        type = "";

                    }
                    else
                    {
                        for (int z = 0; z < modelSplit.Length - 1; z++)
                        {
                            model += modelSplit[z] + " ";
                        }

                        type = modelSplit[modelSplit.Length - 1];
                    }


                    checkType(type, otherAdds);


                    //frameLabelFinalFile = new SLDocument("template.xlsx");
                    //prva vrstica
                    document.SetCellValue(indexBed3, 1, "ORDINE:");
                    document.SetCellValue(indexBed3, 3, bedOrderNumber);
                    document.SetCellValue(indexBed3, 7, bedDeliveryCompany + bedRif);
                    document.SetCellValue(indexBed3, 10, stickerDate);
                    document.SetCellValue(indexBed3, 13, bedOrderNumLocal);
                    indexBed3++;


                    modelTypeSize = model + " " + type + "    " + dimension;
                    if (modelTypeSize.Length > 28)
                    {
                        SLStyle smallerFontNameSize = document.CreateStyle();
                        smallerFontNameSize.Font.FontSize = 11;
                        document.SetCellStyle(indexBed3, 1, smallerFontNameSize);
                    }
                    document.SetCellValue(indexBed3, 1, modelTypeSize);

                    document.SetCellValue(indexBed3, 8, descriptionBed1);
                    indexBed3++;

                    document.SetCellValue(indexBed3, 1, adds);


                    try
                    {
                        Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                        barkoda = bedOrderNumber + "-" + first + "-" + modelbed + "-" + "TESTATA";
                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", "");
                        Image image = brCode.Draw(barkoda, 37); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali
                        image.Save("BedBarCode.gif");                                            // image.Save("frameBarCode.gif");

                        SLPicture pic = new SLPicture("BedBarCode.gif");
                        pic.SetPosition(indexBed3, 1);
                        document.InsertPicture(pic);
                        barkoda = "";

                    }
                    catch
                    {

                    }
                    document.Save();
                    indexBed3++;
                    indexBed3++;
                }
                if (sheet == "5" && specialSticker.Checked)
                {
                    document.SelectWorksheet(sheet);


                    string[] addsSplit;
                    string fabricType;
                    string modelTypeSize;

                    addsSplit = adds.Split(',');
                    fabricType = addsSplit[0].Split(' ')[0];




                    string otherAdds = "";

                    for (int k = 1; k < addsSplit.Length; k++)
                    {
                        otherAdds = addsSplit[k] + " ";
                    }
                    Console.WriteLine(otherAdds);
                    string[] modelSplit = System.Text.RegularExpressions.Regex.Replace(modelbed, @"\s+", " ").Split(' ');

                    string model = "";
                    string type;


                    if (modelSplit.Length == 1)
                    {
                        model = modelSplit[0];
                        type = "";
                        descriptionBed1 = "SIMPLY BED";

                    }
                    else
                    {
                        for (int z = 0; z < modelSplit.Length - 1; z++)
                        {
                            model += modelSplit[z] + " ";
                        }

                        type = modelSplit[modelSplit.Length - 1];
                    }



                    checkType(type, otherAdds);

                    // XSSFFont font = wb.create;




                    //frameLabelFinalFile = new SLDocument("template.xlsx");
                    //prva vrstica
                    //document.SetCellValue(indexBed5, 1, "ORDINE:");
                    //document.SetCellValue(indexBed5, 3, bedOrderNumber);
                    //document.SetCellValue(indexBed5, 7, bedDeliveryCompany + bedRif);
                    document.SetCellValue(indexBed5, 10, stickerDate);
                    document.SetCellValue(indexBed5, 13, bedOrderNumLocal.Split('/')[1]);
                    indexBed5++;


                    modelTypeSize = model + " " + type + "    " + dimension;
                    if (modelTypeSize.Contains("H25"))
                    {
                        document.SetCellValue(indexBed5, 1, "H25     " +type + "   " + dimension);
                    }
                    else
                    {
                        if (modelTypeSize.Contains("COMPACT"))
                        {
                            document.SetCellValue(indexBed5, 1, "COMPACT   " + type + "   " + dimension);
                        }
                        else
                        {
                            document.SetCellValue(indexBed5, 1, type + "   " + dimension);
                        }
                    }



                    /*
                    if (modelTypeSize.Length > 28)
                    {
                        SLStyle smallerFontNameSize = document.CreateStyle();
                        smallerFontNameSize.Font.FontSize = 11;
                        document.SetCellStyle(indexBed5, 1, smallerFontNameSize);
                    }
                    */
                    //document.SetCellValue(indexBed5, 1, type + "   " + dimension);
                    //document.SetCellValue(indexBed5, 1, modelTypeSize);

                    //document.SetCellValue(indexBed5, 8, descriptionBed1);
                    indexBed5++;


                    //POSEBNE OZNAKE ZA KLANFANJE, DA ZNAJO KAJ LANFAT
                    if ((modelTypeSize.Contains("DREAM") || modelTypeSize.Contains("INSIDE") || modelTypeSize.Contains("DIAMOND") || modelTypeSize.Contains("CUBE")) && adds.Contains("FP"))
                    {
                        //adds = adds +"   000000";
                        adds = adds + "   " + (char)127 + (char)127 + (char)127 + (char)127 + (char)127 + (char)127;

                    }

                    if(adds.Contains("STRUCTURAL BASE") || adds.Contains("STRU") || adds.Contains("BASE STR"))
                    {
                        adds = adds + "  STRUCTURAL BASE";
                        document.SetCellValue(indexBed5, 1, "STRUCTURAL BASE");
                    }

                    if(modelTypeSize.Contains("SOMMIER") || modelTypeSize.Contains("VIKY"))
                    {
                        document.SetCellValue(indexBed5, 1, "SOMMIER");
                    }
                    if (modelTypeSize.Contains("NICOLE"))
                    {
                        document.SetCellValue(indexBed5, 1, "NICOLE");
                    }


                    if (adds.Contains("FP") || adds.Contains("PELX") || adds.Contains("PBA"))
                    {
                        document.SetCellValue(indexBed5, 1, adds);
                    }


                    //document.SetCellValue(indexBed5, 1, adds);

                    /*
                    try
                    {
                        Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                        barkoda = bedOrderNumber + "-" + first + "-" + modelbed + "-" + "FUSTO BASE";
                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", " ");
                        Image image = brCode.Draw(barkoda, 37); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali
                        image.Save("BedBarCode.gif");                                            // image.Save("frameBarCode.gif");

                        SLPicture pic = new SLPicture("BedBarCode.gif");
                        pic.SetPosition(indexBed2, 1);
                        document.InsertPicture(pic);
                        //pic.SetPosition(odkod2, 14);
                        //document.InsertPicture(pic);


                        //barCodeIndex = barCodeIndex + 9;


                        barkoda = "";




                    }
                    catch
                    {

                    }
                    */
                    document.Save();
                    indexBed5++;
                    indexBed5++;
                }
            }
            descriptionBed1 = "";
            descriptionBed2 = "";

        }
        public void makeStickerLegs(string sheet, string first, string bedOrderNumber, string deliveyCompany, string bedrif, string modelbed, string adds, DateTime date, int pieces, string bedOrderNumLocal)
        {
            string shrani1 = pathDestopFolder + "\\" + dateToday + "BED.xlsx";

            for (int u = 0; u < pieces; u++)
            {
                document = new SLDocument(shrani1);
                if (sheet == "4")
                {
                    document.SelectWorksheet(sheet);

                    string modelTypeSize;
                    string[] addsSplit;
                    string fabric;
                    string[] fabricSplit;
                    string fabricType;
                    string color;

                    addsSplit = adds.Split(',');
                    fabricType = addsSplit[0].Split(' ')[0];

                    color = addsSplit[0].Split(' ')[1];




                    Console.WriteLine(fabricType + " " + color);

                    string otherAdds = "";

                    for (int k = 1; k < addsSplit.Length; k++)
                    {
                        otherAdds = addsSplit[k] + " ";
                    }
                    Console.WriteLine(otherAdds);

                    string[] modelSplit = System.Text.RegularExpressions.Regex.Replace(modelbed, @"\s+", " ").Split(' ');

                    string model = "";
                    string type;




                    if (modelSplit.Length == 1)
                    {
                        model = modelSplit[0];
                        type = "";
                        descriptionBed1 = "";

                    }
                    else
                    {
                        for (int z = 0; z < modelSplit.Length - 1; z++)
                        {
                            model += modelSplit[z] + " ";
                        }

                        type = modelSplit[modelSplit.Length - 1];
                    }
                    descriptionBed1 = "PIEDI";


                    DateTime today = DateTime.Today;


                    checkType(type, otherAdds);

                    //prva vrstica
                    document.SetCellValue(indexBed4, 1, "ORDINE:");
                    document.SetCellValue(indexBed4, 3, bedOrderNumber);
                    document.SetCellValue(indexBed4, 7, bedDeliveryCompany + bedRif);
                    document.SetCellValue(indexBed4, 10, stickerDate);
                    document.SetCellValue(indexBed4, 13, bedOrderNumLocal);
                    indexBed4++;


                    modelTypeSize = model + " " + type;
                    if (modelTypeSize.Length > 28)
                    {
                        SLStyle smallerFontNameSize = document.CreateStyle();
                        smallerFontNameSize.Font.FontSize = 11;
                        document.SetCellStyle(indexBed4, 1, smallerFontNameSize);
                    }
                    document.SetCellValue(indexBed4, 1, modelTypeSize);

                    document.SetCellValue(indexBed4, 8, descriptionBed1);
                    indexBed4++;

                    document.SetCellValue(indexBed4, 1, adds);

                    try
                    {
                        Zen.Barcode.Code128BarcodeDraw brCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;
                        barkoda = bedOrderNumber + "-" + first + "-" + modelbed + "-" + "PIEDI";
                        barkoda = System.Text.RegularExpressions.Regex.Replace(barkoda, @"\s+", "");
                        Image image = brCode.Draw(barkoda, 37); // številka pomeni višino, širina je odvisna od števila znakov // kako bomo pozicionirali
                        image.Save("BedBarCode.gif");                                            // image.Save("frameBarCode.gif");

                        SLPicture pic = new SLPicture("BedBarCode.gif");
                        pic.SetPosition(indexBed4, 0.5);
                        document.InsertPicture(pic);

                        barkoda = "";

                    }
                    catch
                    {

                    }
                    document.Save();
                    indexBed4++;
                    indexBed4++;
                }
            }
            descriptionBed1 = "";
            descriptionBed2 = "";

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
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

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (checkFileFormat(sender, e, datoteka))
            {

                if (checkBoxSmallSticker.Checked)
                {
                    SLDocument fileNarocila = new SLDocument(datoteka); //open order file
                    Console.WriteLine("ZAČETEK");
                    Console.WriteLine("IME LISTE:" + datoteka);

                    string koncnaPredlogaPathFrame = pathDestopFolder + "\\templateStickers3.xlsx";

                    SLWorksheetStatistics stats = fileNarocila.GetWorksheetStatistics(); // stats for order file, to get last row

                    smallStickerDocument = new SLDocument(koncnaPredlogaPathFrame);

                    string shrani = pathDestopFolder + "\\" + dateToday + "SMALL STICKER.xlsx";

                    MessageBox.Show(shrani);
                    smallStickerDocument.SaveAs(shrani); //save sticker file
                    int smallStickerI = 1;
                    int smallStickerJ = 1;

                    string modelCodeFrame;

                    smallStickerDocument = new SLDocument(shrani);



                    for (int i = 3; i <= stats.NumberOfRows; i++)
                    {

                        //
                        Console.WriteLine(stats.NumberOfRows);
                        Console.WriteLine(i);

                        ordineFrame = fileNarocila.GetCellValueAsString(i, 5); //ORDINE
                        personalizationFrame = fileNarocila.GetCellValueAsString(i, 6);
                        modelFrame = fileNarocila.GetCellValueAsString(i, 7);
                        typeFrame = fileNarocila.GetCellValueAsString(i, 8);
                        modelCodeFrame = fileNarocila.GetCellValueAsString(i, 9);
                        sizeXFrame = fileNarocila.GetCellValueAsString(i, 10);
                        sizeYFrame = fileNarocila.GetCellValueAsString(i, 11);
                        deliveyCompanyFrame = fileNarocila.GetCellValueAsString(i, 16);
                        rifFrame = fileNarocila.GetCellValueAsString(i, 17);
                        descriptionFrame = fileNarocila.GetCellValueAsString(i, 18)+"  "+ fileNarocila.GetCellValueAsString(i, 12) + "  " + fileNarocila.GetCellValueAsString(i, 13);
                        descriptionFrameAddition = fileNarocila.GetCellValueAsString(i, 19) + "  " + fileNarocila.GetCellValueAsString(i, 14) + "  " + fileNarocila.GetCellValueAsString(i, 15);
                        piecesFrame = fileNarocila.GetCellValueAsInt32(i, 20);



                        for (int NumberOfLabels = 1; NumberOfLabels <= piecesFrame; NumberOfLabels++)
                        {
                            smallStickerFrameSpecial(ordineFrame, rifFrame, deliveyCompanyFrame, modelFrame, sizeXFrame, sizeYFrame, descriptionFrame, descriptionFrameAddition, smallStickerI, smallStickerJ);
                            //smallStickerFrame(ordineFrame, rifFrame, modelBed, sizeXFrame, sizeYFrame, column10, descriptionFrameAddition, descriptionFrame, descriptionFrame, smallStickerI, smallStickerJ);
                            if (smallStickerJ < 15)
                            {
                                smallStickerJ += 7;
                            }
                            else
                            {
                                smallStickerJ = 1;
                                smallStickerI += 11;
                            }
                            //smallStickerDocument.Save();
                        }
                        //smallStickerDocument.Save();
                        addsString = "";
                    }
                    smallStickerDocument.Save();
                    fileNarocila.CloseWithoutSaving(); //close order file
                    MessageBox.Show("Nalepke so kreirane."); //messsage shot for successful sticker create
                }
                else
                {
                    SLDocument fileNarocila = new SLDocument(datoteka); //open order file
                    Console.WriteLine("ZAČETEK");
                    Console.WriteLine("IME LISTE:" + datoteka);

                    string koncnaPredlogaPathFrame = pathDestopFolder + "\\templateStickers3.xlsx";

                    SLWorksheetStatistics stats = fileNarocila.GetWorksheetStatistics(); // stats for order file, to get last row

                    smallStickerDocument = new SLDocument(koncnaPredlogaPathFrame);

                    string shrani = pathDestopFolder + "\\" + dateToday + "SMALL STICKER.xlsx";

                    MessageBox.Show(shrani);
                    smallStickerDocument.SaveAs(shrani); //save sticker file
                    int smallStickerI = 1;
                    int smallStickerJ = 1;

                    smallStickerDocument = new SLDocument(shrani);



                    for (int i = 3; i <= stats.NumberOfRows; i++)
                    {

                        //
                        Console.WriteLine(stats.NumberOfRows);
                        Console.WriteLine(i);

                        ordineFrame = fileNarocila.GetCellValueAsString(i, 2); //ORDINE
                        rifFrame = fileNarocila.GetCellValueAsString(i, 3);
                        deliveyCompanyFrame = fileNarocila.GetCellValueAsString(i, 4);
                        modelFrame = fileNarocila.GetCellValueAsString(i, 5);
                        sizeXFrame = fileNarocila.GetCellValueAsString(i, 6);
                        sizeYFrame = fileNarocila.GetCellValueAsString(i, 7);
                        descriptionFrame = fileNarocila.GetCellValueAsString(i, 8);
                        descriptionFrameAddition = fileNarocila.GetCellValueAsString(i, 9);
                        piecesFrame = fileNarocila.GetCellValueAsInt32(i, 10);

                        for (int NumberOfLabels = 1; NumberOfLabels <= piecesFrame; NumberOfLabels++)
                        {
                            smallStickerFrame(ordineFrame, rifFrame, deliveyCompanyFrame, modelFrame, sizeXFrame, sizeYFrame, descriptionFrame, descriptionFrameAddition, smallStickerI, smallStickerJ);
                            //smallStickerFrame(ordineFrame, rifFrame, modelBed, sizeXFrame, sizeYFrame, column10, descriptionFrameAddition, descriptionFrame, descriptionFrame, smallStickerI, smallStickerJ);
                            if (smallStickerJ < 15)
                            {
                                smallStickerJ += 7;
                            }
                            else
                            {
                                smallStickerJ = 1;
                                smallStickerI += 11;
                            }
                            //smallStickerDocument.Save();
                        }
                        //smallStickerDocument.Save();
                        addsString = "";
                    }
                    smallStickerDocument.Save();
                    fileNarocila.CloseWithoutSaving(); //close order file
                    MessageBox.Show("Nalepke so kreirane."); //messsage shot for successful sticker create
                }
            }
        }
    }
}