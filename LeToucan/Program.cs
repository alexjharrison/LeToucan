using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace LeToucan
{
    class LeToucan
    {
        static void Main(string[] args)
        {
            //fail program if incorrect # of inputs from pcdmis found
            if (args.Length != 19)
            {
                Console.WriteLine(args.Length+" parameters detected, require 18\nProgram Failed");
                foreach (string thisArg in args)
                    Console.WriteLine(thisArg);
                Console.ReadKey(false);
                Environment.Exit(0);
            }


            int simultaneousPrints=0; Boolean endProgram=false; double measDiam=0; double measThick=0;
            String shade=""; int matNum=0; double FSDensity=0; String shadeAlias=""; String batchID="";
            int discnum=0; double weight=0; string theoDiam=""; string orderNum=""; string theoThick="";
            string cmmOperator=""; String material=""; double ledgeThick=0; double innerDiam=0; string scaleID="";
            string weightOperator=""; string propertyNum=""; double ledgeOffset=0; double concentricity=0;
            double PSDensity=0; double EF=0; double shrinkage=0;
            string line2 = ""; string[] line2split = { "", "" }; int batchsize = 0;

            try
            {
            //assign args into variables and declare more
            simultaneousPrints = Convert.ToInt32(args[17]); //this is how after how many discs labels are printed
            endProgram = Convert.ToBoolean(args[18]);
            measDiam = Convert.ToDouble(args[11]); measThick = Convert.ToDouble(args[12]); 
            shade = args[6]; matNum = Convert.ToInt32(args[9]); 
            FSDensity = Convert.ToDouble(args[10]); shadeAlias = args[8];
            batchID = args[1]; orderNum = args[2]; 
            discnum = Convert.ToInt32(args[0]); 
            theoDiam = args[7]; theoThick = args[4];
            cmmOperator = Environment.UserName; material = args[5];
            ledgeThick = Math.Round(Convert.ToDouble(args[13]),3); innerDiam = Math.Round(Convert.ToDouble(args[14]),3);
            ledgeOffset = Math.Round(Convert.ToDouble(args[15]),3); concentricity = Math.Round(Convert.ToDouble(args[16]),3);
            
            //to be calculated
            }
            catch
            {
                Console.WriteLine("Your variables are garbage\nProgram will now end");
                Console.ReadKey();
                Environment.Exit(0);
            }


            //RFID file header
            string rfidheader = @"ROHL_ID	""SKAL_FAKT""	""NOMINAL_HOEHE_ALIAS""	""NOMINAL_HOEHE""	""MATR_KUERZEL""	""BARCODE_NUTZDATEN""	""LIEFER_NUMMER""	""DRUCK_NUMMER""	""ETIKETTEN_HOEHE""	""SINTER_NR""	""DURCHMESSER""	""QUALITAET""	""LIEFERANT""	""LOT""	""DATUM_EINGANG""	""AUTH_CODE""	""BARCODE""	""ERSATZ_CODE""	""MAT_NAME""	""QUALT""	";

            //declare locations of files
            string weightlocation = "G:/LabX_Export/" + batchID + "_" + orderNum;
            string RFIDLocation = "G:/Equipment/GiS-Topex RFID Printer/Generated RFID Files/" +shade+"/"+ batchID + "/";

            //make temporary copy
            bool fileLocked = true;
            while (fileLocked == true)
            {
                try
                {
                    File.Copy(weightlocation + ".csv", weightlocation + "_tempcopy" + ".csv", true);
                    FileInfo fileInfo = new FileInfo(weightlocation + "_tempcopy" + ".csv");
                    fileInfo.IsReadOnly = false; //if weights file is read only, the copy won't be 
                    fileLocked = false;
                }
                catch
                {
                    ConsoleKeyInfo keyStroke;
                    Console.WriteLine(weightlocation + ".csv does not exist or is locked");
                    Console.WriteLine("Please close the file");
                    Console.WriteLine("Press q to quit or any key to retry\n\n");
                    keyStroke = Console.ReadKey(false);
                    if (keyStroke.KeyChar == 'q') Environment.Exit(0);
                }
            }

            //Find batch size from weights file
            try
            {
                line2 = File.ReadLines(weightlocation + "_tempcopy" + ".csv").Skip(1).Take(1).First();
                line2split = line2.Split(',');
                //Console.WriteLine(line2split.Length);
                batchsize = Convert.ToInt32(line2split[4]);
            }
            catch
            {
                Console.WriteLine("Could not find batch size");
                Console.WriteLine("Weights file is improperly formatted");
                Console.WriteLine("Program will now end\nPress any key");
                Console.ReadKey();
                File.Delete(weightlocation + "_tempcopy" + ".csv");
                Environment.Exit(0);
            }
            
            //on last run do housekeeping
            if (endProgram == true) 
            {
                int samplesInWeightsFile = 0;
                int samplesInRFIDFile = 0;
                try
                {
                    samplesInWeightsFile = CountLinesInFile(weightlocation + "_tempcopy" + ".csv") - 3;
                    samplesInRFIDFile = CountLinesInFile(RFIDLocation + batchID + "_rfid_complete.csv") - 1;
                }
                catch
                {
                    Console.WriteLine(weightlocation + "_tempcopy" + ".csv or " + RFIDLocation + batchID + "_rfid_complete.csv\nis unreadable");
                    File.Delete(weightlocation + "_tempcopy" + ".csv");
                    Console.WriteLine("Program will now end");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                if ((batchsize == samplesInWeightsFile)&&(batchsize==samplesInRFIDFile))
                {
                    try
                    {
                        File.Copy(weightlocation + "_appending" + ".csv", weightlocation + ".csv", true);
                        File.Delete(weightlocation + "_appending" + ".csv");
                        File.Delete(weightlocation + "_tempcopy" + ".csv");
                    }
                    catch
                    {
                        Console.WriteLine(weightlocation + "_appending" + ".csv\nor");
                        Console.WriteLine(weightlocation + "_tempcopy" + ".csv");
                        Console.WriteLine("missing or write protected");
                        Console.WriteLine("Press any key to try again or q to quit...\n");
                        ConsoleKeyInfo keyStroke = Console.ReadKey(false);
                        if (keyStroke.KeyChar == 'q') 
                            Environment.Exit(0);
                    }

                    //Uncomment this when testing for LabsQ
                    //File.Copy(weightlocation + ".csv", @"G:\LabsQ\LabsQ Instrument Link\Import\" + batchID + "_" + orderNum + ".csv");

                }

                else
                {
                    Console.WriteLine("Final Formatting Failed\nNumber of samples does not match batch size");
                    Console.WriteLine("Check " + weightlocation + "_appending.csv\nand\n" + RFIDLocation + batchID + "_rfid_complete.csv\nfor discrepancies");
                    File.Delete(weightlocation + "_tempcopy" + ".csv");
                    Console.WriteLine("Program will now end");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    Environment.Exit(0);
                }


                
                
                Environment.Exit(0);
            }
            
            
                
            
            
            //declare variables for loading weights file columns
            var column1 = new List<string>();
            var column2 = new List<string>();
            var column3 = new List<string>();
            var column4 = new List<string>();
            var column5 = new List<string>();
            var column6 = new List<string>();
            var column7 = new List<string>();
            var column8 = new List<string>();
            var column9 = new List<string>();

            

            //7 columns if still weighing, 9 columns if finished
            //load in weight file columns into list arrays
            if (line2split.Length == 9)
            {
                try
                {
                    using (var rd = new StreamReader(weightlocation + "_tempcopy" + ".csv"))
                    {
                        while (!rd.EndOfStream)
                        {
                            string theNextLine = rd.ReadLine();
                            if (theNextLine == "") break;
                            var splits = theNextLine.Split(',');
                            column1.Add(splits[0]);
                            column2.Add(splits[1]);
                            column3.Add(splits[2]);
                            column4.Add(splits[3]);
                            column5.Add(splits[4]);
                            column6.Add(splits[5]);
                            column7.Add(splits[6]);
                            column8.Add(splits[7]);
                            column9.Add(splits[8]);
                        }
                    }
                }
                catch
                {
                    Console.WriteLine(weightlocation + ".csv" + " has the incorrect number of columns");
                    Console.WriteLine("Program will now end");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
                try 
                {
                //remove blank cells from column 9
                column9.RemoveAll(string.IsNullOrWhiteSpace);
                weight = Convert.ToDouble(column9[discnum]);
                scaleID = column6[1];
                weightOperator = column3[1];
                propertyNum = column7[0];
                }
                catch
                {
                    Console.WriteLine(weightlocation + ".csv" + " has data in the incorrect location");
                    Console.WriteLine("Program will now end");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    File.Delete(weightlocation + "_tempcopy" + ".csv");
                    Environment.Exit(0);
                }
            }
            else
            {
                try
                {
                    using (var rd = new StreamReader(weightlocation + "_tempcopy" + ".csv"))
                    {
                        while (!rd.EndOfStream)
                        {
                            string theNextLine = rd.ReadLine();
                            if (theNextLine == "") break;
                            var splits = theNextLine.Split(',');
                            column1.Add(splits[0]);
                            column2.Add(splits[1]);
                            column3.Add(splits[2]);
                            column4.Add(splits[3]);
                            column5.Add(splits[4]);
                            column6.Add(splits[5]);
                            column7.Add(splits[6]);
                        }
                    }
                }
                catch
                {
                    Console.WriteLine(weightlocation + ".csv" + " has the incorrect number of columns");
                    Console.WriteLine("Program will now end");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    File.Delete(weightlocation + "_tempcopy" + ".csv");
                    Environment.Exit(0);
                }
                try
                {
                    //remove blank cells from column 7
                    column7.RemoveAll(string.IsNullOrWhiteSpace);
                    weight = Convert.ToDouble(column7[discnum]);
                    scaleID = column6[3];
                    weightOperator = column3[1];
                    propertyNum = column7[0];
                }
                catch
                {
                    Console.WriteLine(weightlocation + ".csv" + " has data in the incorrect location");
                    Console.WriteLine("Program will now end");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    File.Delete(weightlocation + "_tempcopy" + ".csv");
                    Environment.Exit(0);
                }

            }
            
            
            File.Delete(weightlocation + "_tempcopy" + ".csv");
           
            Console.WriteLine(scaleID + " " + weightOperator + " " + propertyNum + " " + cmmOperator);

            var finalColumn1 = new List<string>();
            var finalColumn2 = new List<string>();
            var finalColumn3 = new List<string>();
            var finalColumn4 = new List<string>();
            var finalColumn5 = new List<string>();
            var finalColumn6 = new List<string>();
            var finalColumn7 = new List<string>();
            var finalColumn8 = new List<string>();
            var finalColumn9 = new List<string>();
            var finalColumn10 = new List<string>();
            var finalColumn11 = new List<string>();
            //if first disc create header to file
            if(discnum==1)
            {
                //create rfid file header
                Directory.CreateDirectory(RFIDLocation);
                File.WriteAllText(RFIDLocation + batchID+"_rfid_complete" + ".csv", rfidheader );

                finalColumn1 = new List<string>() { "Batch#", batchID, "", "User Name", "Instrument ID#", "Test Method#", "Property#", "Sample#" };
                finalColumn2 = new List<string>() { "Production Order#", orderNum, "", weightOperator, scaleID, "TM700", propertyNum, "Disc Weight" };
                finalColumn3 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Diameter" };
                finalColumn4 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Thickness" };
                finalColumn5 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Density" };
                if(material!="ZTG"&&material!="ZG")
                {
                    finalColumn6 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "EF" };
                    finalColumn7 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Shrinkage" };
                    if(material!="ZFC")
                    {
                        finalColumn8 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Inner Diameter" };
                        finalColumn9 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Ledge Thickness" };
                        finalColumn10 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Ledge Offset" };
                        finalColumn11 = new List<string>() { "", "", "", cmmOperator, "TE-47", "TM700", propertyNum, "Ledge Concentricity" };
                    }
                    
                }
                
                var csv = new StringBuilder();
                for (int i = 0; i <= 7; i++)
                {
                    string newLine;
                    if (material=="ZTG"||material=="ZG")
                    {
                        newLine = string.Join(",", finalColumn1[i], finalColumn2[i], finalColumn3[i], finalColumn4[i], finalColumn5[i]);
                    }
                    else if(material=="ZFC")
                    {
                        newLine = string.Join(",", finalColumn1[i], finalColumn2[i], finalColumn3[i], finalColumn4[i], finalColumn5[i], finalColumn6[i], finalColumn7[i]);
                    }
                    else
                    {
                        newLine = string.Join(",", finalColumn1[i], finalColumn2[i], finalColumn3[i], finalColumn4[i], finalColumn5[i], finalColumn6[i], finalColumn7[i], finalColumn8[i], finalColumn9[i], finalColumn10[i], finalColumn11[i]);
                    }
                    
                    csv.AppendLine(newLine);
                }
                File.WriteAllText(weightlocation + "_appending.csv", csv.ToString());
            }

            PSDensity = CalcDensity(measDiam, measThick, weight, material);
            //Append data for this disc
            string nextLine = "";
            if (material=="ZTG"||material=="ZG")
            {
                nextLine = (discnum + "," + weight + "," + measDiam + "," + measThick + "," + Math.Round(PSDensity,3)  + Environment.NewLine);
                File.AppendAllText(weightlocation + "_appending" + ".csv", nextLine);
            }
            else if (material=="ZFC")
            {
                EF = CalcEF(FSDensity, PSDensity);
                shrinkage = CalcShrink(EF);
                nextLine = (discnum + "," + weight + "," + measDiam + "," + measThick + "," + Math.Round(PSDensity,3) + "," + Math.Round(EF,4) + "," + Math.Round(shrinkage,3) + Environment.NewLine);
                File.AppendAllText(weightlocation + "_appending" + ".csv", nextLine);
                string finalExport = GenerateRFIDOutput(discnum,batchID,theoThick,matNum,material,theoDiam,EF,simultaneousPrints,batchsize,RFIDLocation,shade,shadeAlias,measDiam,rfidheader,shrinkage);
            }
            else
            {
                EF = CalcEF(FSDensity, PSDensity);
                shrinkage = CalcShrink(EF);
                nextLine = (discnum + "," + weight + "," + measDiam + "," + measThick + "," + Math.Round(PSDensity,3) + "," + Math.Round(EF,4) + "," + Math.Round(shrinkage,3)+","+innerDiam+","+ledgeThick+","+ledgeOffset+","+concentricity + Environment.NewLine);
                File.AppendAllText(weightlocation + "_appending" + ".csv", nextLine);
                string finalExport = GenerateRFIDOutput(discnum,batchID,theoThick,matNum,material,theoDiam,EF,simultaneousPrints,batchsize,RFIDLocation,shade,shadeAlias,measDiam,rfidheader,shrinkage);
            }
        }



        static string Encode(long decimalNumber)
        {
            const int radix = 35;
            const int BitsInLong = 64;
            const string Digits = "0123456789abcdefghijkmnopqrstuvwxyz";
            if (radix < 2 || radix > Digits.Length)
                throw new ArgumentException("The radix must be >= 2 and <= " + Digits.Length.ToString());
            if (decimalNumber == 0)
                return "0";
            int index = BitsInLong - 1;
            long currentNumber = Math.Abs(decimalNumber);
            char[] charArray = new char[BitsInLong];
            while (currentNumber != 0)
            {
                int remainder = (int)(currentNumber % radix);
                charArray[index--] = Digits[remainder];
                currentNumber = currentNumber / radix;
            }
            string result = new String(charArray, index + 1, BitsInLong - index - 1);
            if (decimalNumber < 0)
            {
                result = "-" + result;
            }
            return result;
        }
        static double CalcDensity(double diam, double thick, double weight, String material)
        {
            double density = 0;
            if (material == "ZFC")  //Zirlux LGS723a
                density = weight / ((Math.Pow(0.5 * diam, 2.0) * Math.PI) * (thick / 1000));
            else if (material == "ZTR")  //Wieland Translucent Light, Medium... LGS723b
                density = (weight * 1000) / (Math.PI * Math.Pow(diam / 2, 2.0) * thick - (41.878 * 2 / 100.2 * diam));
            else if (material == "ZMO")  //Wieland MO LGS723d
                density = (weight * 1000) / (Math.PI * Math.Pow((diam / 2), 2.0) * thick - (0.4 * Math.PI * diam));
            else if (material == "ZT")  //Wieland T0,T1...  LGS723e
                density = weight * 0.9955 / ((Math.PI * Math.Pow(diam, 2.0) * thick / 4 / 1000 - 2 * 2 * Math.PI * (diam / 2 / 1000) * (0.27 / 2)));
            else if (material == "ZTG")  //Wieland Green Light, Medium... LGS723c
                density = (weight*1000)/((Math.PI*(Math.Pow((diam/2),2.0)*thick)-(41.878*2/100.2*diam)));
            else if (material == "ZG")  //Wieland Green T0,T1,T2... LGS723f
                density = weight * 0.9955 / ((Math.PI * Math.Pow(diam,2) * thick / 4 / 1000 - 2 * 2 * Math.PI * (diam / 2 / 1000) * (0.27 / 2)));
            return density;
        }
        static double CalcEF(double FSDensity,double PSDensity)
        {
            double EF;
            EF = Math.Pow(FSDensity / PSDensity, (1.0 / 3.0));
            return EF;
        }   
        static double CalcShrink(double EF)
        {
            double shrinkage;
            shrinkage = ((1-(1/EF))*100);
            return shrinkage;
        }
        protected static void IsFileLocked(FileInfo file, string filenamepath)
        {
            ConsoleKeyInfo keyStroke;
            bool locked = true;
            FileStream stream = null;
            while (locked==true)
            {
                try
                {
                    stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    locked = false;
                }
                catch (IOException) 
                {
                    //the file is unavailable because it is:
                    //still being written to
                    //or being processed by another thread
                    //or does not exist (has already been processed)
                    Console.WriteLine("File is in use or does not exist");
                    Console.WriteLine("Close the file");
                    Console.WriteLine(filenamepath);
                    Console.WriteLine("Press any key to try again or q to quit...\n");
                    keyStroke = Console.ReadKey(false);
                    if (keyStroke.KeyChar == 'q') Environment.Exit(0);
                }
            }
            /*finally
            {
                if (stream != null)
                    stream.Close();
            }
            //file is not locked
            return false;*/
        }
        static String GenerateWielandID(string batchID)
        {
            var IVMI = new List<string>();
            var Wieland = new List<string>();
            int lastWielandNum=0;
            try
            {
                using (var rd = new StreamReader("G:/LabX_Export/LabsQ_LabX_Integration/WielandID Database.csv"))
                {
                    while (!rd.EndOfStream)
                    {
                        string theNextLine = rd.ReadLine();
                        if (theNextLine == "") break;
                        var splits = theNextLine.Split(',');
                        IVMI.Add(splits[0]);
                        Wieland.Add(splits[1]);
                    }
                }
            }
            catch
            {
                Console.WriteLine("G:/LabX_Export/LabsQ_LabX_Integration/WielandID Database.csv");
                Console.WriteLine("Program will now end");
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                Environment.Exit(0);
            }

            for (int i = 0; i < IVMI.Count; i++)
            {
                if (string.Equals(IVMI[i], batchID)) return Wieland[i];
                lastWielandNum = Convert.ToInt32(Wieland[i]);
            }
            IVMI.Add(batchID); Wieland.Add(Convert.ToString(lastWielandNum + 1));
            var csv = new StringBuilder();
            string newLine = "";
            for (int i = 0; i < IVMI.Count;i++ )
            {
                newLine = string.Join(",", IVMI[i], Wieland[i]);
                csv.AppendLine(newLine);
            }
            File.WriteAllText("G:/LabX_Export/LabsQ_LabX_Integration/WielandID Database.csv",csv.ToString());
            return Convert.ToString(lastWielandNum + 1);
        }
        static String GenerateRFIDOutput(int discnum, string batchID, string theoThick, int matNum, string material, string theoDiam, double EF, int simultaneousPrints, int batchSize, string RFIDLocation, string shade, string shadeAlias,double measDiam, string RFIDHeader, double shrinkage)
        {

            //generate rfid output csv
            string efDigits; string discID; string encodedDiscNum; string discInfoEncoded; string barcodeNutzdaten; string barcode; string encodedCode; string finalExport; int heightAlias; string WielandID; string stringNum; string shadeAlias2="";
            if (material == "ZFC")
            {
                heightAlias = 0; matNum = 0;
                WielandID = batchID.Substring(1, 5);
                shadeAlias = shade; shadeAlias2 = shade;
                        
            }
            else
            {
                heightAlias = (Convert.ToInt32(theoThick) * 4 - 19);
                WielandID = GenerateWielandID(batchID);
            }
            if (discnum < 10) stringNum = String.Concat("00", Convert.ToString(discnum));
            else if (discnum < 100) stringNum = String.Concat("0", Convert.ToString(discnum));
            else stringNum = String.Concat(Convert.ToString(discnum));
            discID = String.Concat(WielandID, stringNum);
            
            efDigits = Convert.ToString(Convert.ToString((Math.Round(EF, 4) - 1) * 10000));
            encodedDiscNum = Encode(Convert.ToInt32(discID));
            discInfoEncoded = Convert.ToString(Encode(Convert.ToInt32(String.Concat(Convert.ToString(matNum), Convert.ToString(Math.Abs(Convert.ToInt32(efDigits))), Convert.ToString(heightAlias)))));
            barcodeNutzdaten = String.Concat(discID, Convert.ToString(matNum), efDigits, Convert.ToString(heightAlias));
            barcode = String.Concat(barcodeNutzdaten, "4418753237");
            encodedCode = String.Concat(Convert.ToString(encodedDiscNum), "-", discInfoEncoded, "-", "2e4mahh");
            
            //export rfid to ..._complete.csv 
            if(material=="ZFC")
                finalExport = String.Concat(Environment.NewLine, discID, "\t", efDigits, "\t", theoDiam, "\t", theoThick, "\t", matNum, "\t", barcodeNutzdaten, "\t", WielandID, "\t", stringNum, "\t\t", batchID, "\t", Math.Round(shrinkage, 3), "\tH\t\t", batchID, "\t\t4418753237\t", barcode, "\t", encodedCode, "\t", shadeAlias, "\t"+shadeAlias2+"\t");
            else 
                finalExport = String.Concat(Environment.NewLine, discID, "\t", efDigits, "\t", heightAlias, "\t", theoThick, ".00\t", matNum, "\t", barcodeNutzdaten, "\t", batchID, "\t", stringNum, "\t\t", batchID, "\t", Math.Round(measDiam, 2), "\tH\t\t", batchID, "\t\t4418753237\t", barcode, "\t", encodedCode, "\t", shadeAlias, "\t\t");

            File.AppendAllText(RFIDLocation + batchID + "_rfid_complete.csv", finalExport);

            //export rfid to smaller sheet
            if(simultaneousPrints==1)
            {
                File.WriteAllText(RFIDLocation + batchID + "_rfid_" + discnum + ".csv", RFIDHeader);
                File.AppendAllText(RFIDLocation + batchID + "_rfid_" + discnum + ".csv", finalExport);
                {
                    //put in command to start rfid label printer
                    if (material == "ZFC") //Zirlux
                        System.Diagnostics.Process.Start("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \""+RFIDLocation + batchID + "_rfid_" + discnum + ".csv \"  /start /hidden");
                    
                    else  //Wieland
                        System.Diagnostics.Process.Start("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Topex_Printer\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv \"  /start /hidden");
                    Console.ReadKey();
                }
            }
            else
            {
                int highNum = 0;
                int lowNum;
                for (int i = 0; highNum < discnum; i++)
                    highNum = simultaneousPrints * i;
                lowNum = highNum - simultaneousPrints + 1;
                if (discnum % simultaneousPrints == 1)
                    File.WriteAllText(RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv", RFIDHeader);
                File.AppendAllText(RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv", finalExport);
                if ((discnum % simultaneousPrints == 0) || (discnum == batchSize))
                {
                    //put in command to start rfid label printer
                    if (material == "ZFC") //Zirlux
                    {

                        Console.WriteLine("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv \"  /start /hidden");
                        System.Diagnostics.Process.Start("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv \"  /start /hidden");
                    }
                    else  //Wieland
                        System.Diagnostics.Process.Start("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Topex_Printer\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv \"  /start /hidden");
                    Console.ReadKey();
                }
            }
            return finalExport;
        }
        static int CountLinesInFile(string f)
        {
            int count = 0;
            using (StreamReader r = new StreamReader(f))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    if (line == "") break; 
                    count++;
                }
            }
            return count;
        }
    }
}
