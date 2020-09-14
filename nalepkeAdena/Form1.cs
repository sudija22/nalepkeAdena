﻿using System;
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
        string datoteka = null;
        string rif;
        string model;
        string vrsta; //1,2,3
        string dolzina;
        string sirina;
        int steviloKosov;
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
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                datoteka = openFileDialog1.FileName;
                try
                {
                    string text = File.ReadAllText(datoteka);
                    size = text.Length;
                }
                catch (IOException)
                {
                }
            }
        }

        private void potrditev_Click(object sender, EventArgs e)
        {
            if (datoteka == null)
            {
                MessageBox.Show("Najprej izberi datoteko");
            }
            else
            {
                string formatCheck = datoteka.Substring((datoteka.Length - 4), 4);
                if (formatCheck != "xlsx" && formatCheck != ".ods")
                {
                    MessageBox.Show("Format ni pravilen. Pravilen format je '.xlsx'");
                }
                else
                {
                    SLDocument fileNarocila = new SLDocument(datoteka); //open order file

                    //SLDocument fileNalepke = new SLDocument("predloga.xlsx");// open template file
                    string pathPredloga = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string kocnoPredlogaPath = pathPredloga + "\\nalepkeProgram\\predloga.xlsx";
                    SLDocument fileNalepke = new SLDocument(kocnoPredlogaPath);
                    //SLDocument fileNalepkeOblecene = new SLDocument("predlogaOblecene.xlsx");
                    SLStyle fontMereLength = fileNalepke.CreateStyle();
                    SLStyle fontModelLength1 = fileNalepke.CreateStyle();
                    SLStyle odebeljeno = fileNalepke.CreateStyle();

                    fontMereLength.Font.FontSize = 14; // change if string lengt of size is too long 9 chars.
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
                    formatiranDatum += leto;


                    //MessageBox.Show(steviloStrani.ToString());
                    SLWorksheetStatistics stats = fileNarocila.GetWorksheetStatistics(); // stats for order file, to get last row
                    for (int i = 3; i <= stats.EndRowIndex; i++)
                    {
                        vrsticaCheck = fileNarocila.GetCellValueAsString(i, 2);
                        if (vrsticaCheck != "")
                        {
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

                        }
                    }

                    string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //get current user destop path
                    string shrani = pathPredloga + "\\nalepkeProgram\\" + datum + " NALEPKE.xlsx";
                    //string shrani = path + "\\" + datum + "NALEPKE.xlsx"; // format save name of file to save on user destop
                    //MessageBox.Show(shrani);
                    fileNalepke.SaveAs(shrani); //save sticker file
                    fileNarocila.CloseWithoutSaving(); //close order file
                    MessageBox.Show("Nalepke so kreirane."); //messsage shot for successful sticker create


                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}

//time 18,5 h +2
// todo list:
// bonsai salus string length   
//AUGO vse v isto polje. 
// kaj je z innovo 3/2 ???
//no pistoni v rifu ??        
