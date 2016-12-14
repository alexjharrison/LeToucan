using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
namespace LeToucan
{
    class LeToucan
    {
        public static string thingOuttaSpec = "";
        public static string isConforming = "Conform";
        public static string whatsOuttaSpec = "";
        static void Main(string[] args)
        {
            //fail program if incorrect # of inputs from pcdmis found
            if (args.Length != 24) 
            {
                Console.WriteLine(args.Length+" parameters detected, require 24\nProgram Failed");
                foreach (string thisArg in args)
                    Console.WriteLine(thisArg);
                Console.ReadKey(false);
                Environment.Exit(0);
            }

            /* Arg List
             * 1.Disc Number    2.BatchID       3.OrderNum
             * 4.Material Num   5.Theo Thick    6.MaterialCode
             * 7.Shade          8.Theo Diam     9.ShadeAlias
             * 10.Material Code 11.FS Density   12.Meas Diam
             * 13.Meas Thick    14.Ledge Thick  15.Inner Diam
             * 16.Ledge Offset  17.Concentricity18.Simultaneous Prints
             * 19.End Program Bool Trigger      20.Print Toggler
             * 21. Full Description for Label   22.Packaged Material Number
             * 23. CMM ID       24.Total Batch Quantity
             */

            int simultaneousPrints=0; Boolean endProgram=false; double measDiam=0; double measThick=0;
            String shade=""; int matNum=0; double FSDensity=0; String shadeAlias=""; String batchID="";
            int discnum=0; double weight=0; string theoDiam=""; string orderNum=""; string theoThick="";
            string cmmOperator=""; String material=""; double ledgeThick=0; double innerDiam=0; string scaleID="";
            string weightOperator=""; string propertyNum=""; double ledgeOffset=0; double concentricity=0;
            double PSDensity=0; double EF=0; double shrinkage=0;
            string line2 = ""; string[] line2split = { "", "" }; int batchsize = 0;
            bool printToggle = true; string materialID = ""; string fullMaterialID = ""; string packagedMatNum = "";
            string CMMID = ""; int totalQuantity = 0;


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
                ledgeThick = Convert.ToDouble(args[13]); innerDiam = Convert.ToDouble(args[14]);
                ledgeOffset = Convert.ToDouble(args[15]); concentricity = Convert.ToDouble(args[16]);
                printToggle = Convert.ToBoolean(args[19]); materialID = args[3];
                fullMaterialID = args[20]; packagedMatNum = args[21]; CMMID = args[22];
                totalQuantity = Convert.ToInt32(args[23]);
            
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
            
            //old cleanup location
            
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
            
            //new cleanup location
            //on last run do housekeeping
            if (endProgram == true)
            {
                //added code to reprint last labels
                if ((discnum != batchsize) && (material != "ZTG") && (material != "ZG"))
                {
                    PrintBoxLabels(discnum, theoThick, matNum, batchID, material, materialID, batchsize, @"G:\Prod_Labels\Zirconia\Box Labels\", 1, shade, fullMaterialID, packagedMatNum);

                    bool retain2;
                    PSDensity = CalcDensity(measDiam, measThick, weight, material, innerDiam, ledgeThick);
                    EF = CalcEF(FSDensity, PSDensity);
                    shrinkage = CalcShrink(EF);
                    retain2 = SpecCheck(material, matNum, measDiam, measThick, PSDensity, innerDiam, ledgeThick, ledgeOffset, concentricity, EF, materialID);
                    ProcessStartInfo theDruck = new ProcessStartInfo();
                    theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";
                    theDruck.CreateNoWindow = true;
                    theDruck.UseShellExecute = false;
                    theDruck.WindowStyle = ProcessWindowStyle.Hidden;
                    //zirlux retain
                    if ((material == "ZFC") && retain2)
                        theDruck.Arguments = " /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring_Retain.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                    //zirlux non-retain
                    else if ((material == "ZFC") && !retain2)
                        theDruck.Arguments = " /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                    //new wieland retain (ZMU,ZCMO,ZCLT)
                    else if ((material == "ZMU" || material == "ZCMO" || material == "ZCLT") && retain2)
                        theDruck.Arguments = " /P \"G:\\Topex_Printer\\Ivoclar_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                    //new wieland non-retain (ZMU,ZCMO,ZCLT)
                    else if ((material == "ZMU" || material == "ZCMO" || material == "ZCLT") && !retain2)
                        theDruck.Arguments = " /P \"G:\\Topex_Printer\\Ivoclar_CE0123_retain.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                    //old wieland retain (ZMO,ZT,ZTR)
                    else if (retain2)
                        theDruck.Arguments = " /P \"G:\\Topex_Printer\\Wieland Label Templates\\Wieland_CE0123_Retain.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                    //old wieland non-retain(ZMO,ZT,ZTR)
                    else if (!retain2)
                        theDruck.Arguments = " /P \"G:\\Topex_Printer\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                    

                    using (Process exeProcess = Process.Start(theDruck))
                    {
                        exeProcess.WaitForExit(20000);
                    }
                }
                //end added code

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
                /*
                if ((batchsize == samplesInWeightsFile)&&(batchsize==samplesInRFIDFile))
                {
                 * */
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
                /*
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
                */
                Environment.Exit(0);

            }

            
            File.Delete(weightlocation + "_tempcopy" + ".csv");
           

            var finalColumn1 = new List<string>();
            var finalColumn2 = new List<string>();
            var finalColumn3 = new List<string>();
            var finalColumn4 = new List<string>();
            var finalColumn5 = new List<string>();
            var finalColumn6 = new List<string>();
            var finalColumn7 = new List<string>();
            var finalColumn8 = new List<string>();  
            var finalColumn9 = new List<string>();
            //var finalColumn10 = new List<string>();
            //var finalColumn11 = new List<string>();
            //if first disc create header to file
            if(discnum==1)
            {
                //create rfid file header
                Directory.CreateDirectory(RFIDLocation);
                File.WriteAllText(RFIDLocation + batchID+"_rfid_complete" + ".csv", rfidheader );
                DateTime thisDay = DateTime.Today;


                if(material=="ZG"||material=="ZTG")
                {
                    finalColumn1 = new List<string>() { "Batch#", batchID, "", "Sample#" };
                    finalColumn2 = new List<string>() { "Production Order#", orderNum, "", "Disc Weight" };
                    finalColumn3 = new List<string>() { "Start Date", thisDay.ToString("d"), "", "Diameter" };
                    finalColumn4 = new List<string>() { "Disc Thickness", theoThick, "", "Thickness" };
                    finalColumn5 = new List<string>() { "Quantity", totalQuantity.ToString(), "", "Density" };
                    finalColumn6 = new List<string>() { "Disc Shade", shade, "", "" };
                    finalColumn7 = new List<string>() { "Theoretical Density", FSDensity.ToString(), "", "" };
                    finalColumn8 = new List<string>() { "CMM IVMI #", CMMID, "", "" };
                    finalColumn9 = new List<string>() { "", "", "", "" };
                }
                else
                {
                    finalColumn1 = new List<string>() { "Batch#", batchID, "", "Sample#" };
                    finalColumn2 = new List<string>() { "Production Order#", orderNum, "", "Disc Weight" };
                    finalColumn3 = new List<string>() { "Start Date", thisDay.ToString("d"), "", "Diameter" };
                    finalColumn4 = new List<string>() { "Disc Thickness", theoThick, "", "Thickness" };
                    finalColumn5 = new List<string>() { "Quantity", totalQuantity.ToString(), "", "Density" };
                    finalColumn6 = new List<string>() { "Disc Shade", shade, "", "EF" };
                    finalColumn7 = new List<string>() { "Theoretical Density", FSDensity.ToString(), "", "Shrinkage" };
                    finalColumn8 = new List<string>() { "CMM IVMI #", CMMID, "", "Status" };
                    finalColumn9 = new List<string>() { "", "", "", "Nonconform Items" };
                }
                
                    
                    
                
                
                var csv = new StringBuilder();
                for (int i = 0; i < 4; i++)
                {
                    string newLine;
                    newLine = string.Join(",", finalColumn1[i], finalColumn2[i], finalColumn3[i], finalColumn4[i], finalColumn5[i], finalColumn6[i], finalColumn7[i], finalColumn8[i], finalColumn9[i]);
                    csv.AppendLine(newLine);
                }
                File.WriteAllText(weightlocation + "_appending.csv", csv.ToString());
            }

            PSDensity = CalcDensity(measDiam, measThick, weight, material,innerDiam,ledgeThick);
            //Append data for this disc
            string nextLine = "";
            bool retain;
            if (material=="ZTG"||material=="ZG")
            {
                nextLine = (discnum + "," + weight + "," + measDiam + "," + measThick + "," + Math.Round(PSDensity, 3)  + Environment.NewLine);
                File.AppendAllText(weightlocation + "_appending" + ".csv", nextLine);
            }
            else if (material=="ZFC")
            {
                EF = CalcEF(FSDensity, PSDensity);
                shrinkage = CalcShrink(EF);
                retain = SpecCheck(material, matNum, measDiam, measThick, PSDensity, innerDiam, ledgeThick, ledgeOffset, concentricity, EF, materialID);
                nextLine = (discnum + "," + weight + "," + measDiam + "," + measThick + "," + Math.Round(PSDensity, 3) + "," + Math.Round(EF, 4) + "," + Math.Round(shrinkage, 3) + "," + isConforming + "," + whatsOuttaSpec + Environment.NewLine);
                File.AppendAllText(weightlocation + "_appending" + ".csv", nextLine);
                string finalExport = GenerateRFIDOutput(discnum,batchID,theoThick,matNum,material,theoDiam,EF,simultaneousPrints,batchsize,RFIDLocation,shade,shadeAlias,measDiam,rfidheader,shrinkage,retain,printToggle);
            }
            else
            {
                EF = CalcEF(FSDensity, PSDensity);
                shrinkage = CalcShrink(EF);
                retain = SpecCheck(material, matNum, measDiam, measThick, PSDensity, innerDiam, ledgeThick, ledgeOffset, concentricity, EF, materialID);
                nextLine = (discnum + "," + weight + "," + measDiam + "," + measThick + "," + Math.Round(PSDensity, 3) + "," + Math.Round(EF, 4) + "," + Math.Round(shrinkage, 3) + "," + isConforming + "," + whatsOuttaSpec + Environment.NewLine);
                //nextLine = (discnum + "," + weight + "," + measDiam + "," + measThick + "," + Math.Round(PSDensity,3) + "," + Math.Round(EF,4) + "," + Math.Round(shrinkage,3)+","+Math.Round(innerDiam,2)+","+Math.Round(ledgeThick,2)+","+Math.Round(ledgeOffset,2)+","+Math.Round(concentricity,1) + Environment.NewLine);
                File.AppendAllText(weightlocation + "_appending" + ".csv", nextLine);
                string finalExport = GenerateRFIDOutput(discnum,batchID,theoThick,matNum,material,theoDiam,EF,simultaneousPrints,batchsize,RFIDLocation,shade,shadeAlias,measDiam,rfidheader,shrinkage,retain,printToggle);
            }
            
            string labelTemplate = @"G:\Equipment\GiS-Topex RFID Printer\Datamax Template Files\";
            
            if(printToggle)
            {
                if (material == "ZFC" || material == "ZMO" || material == "ZTR" || material == "ZT")
                {
                    if (discnum == 1 || discnum == batchsize)
                        PrintBoxLabels(discnum, theoThick, matNum, batchID, material, materialID, batchsize, labelTemplate, 2, shade, fullMaterialID,packagedMatNum);
                    else
                        PrintBoxLabels(discnum, theoThick, matNum, batchID, material, materialID, batchsize, labelTemplate, 1, shade, fullMaterialID,packagedMatNum);
                }
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
        static double CalcDensity(double diam, double thick, double weight, String material, double innerDiam, double ledgeThick )
        {
            double density = 0;
            if (material == "ZFC")  //Zirlux LGS723a
                density = weight / ((Math.Pow(0.5 * diam, 2.0) * Math.PI) * (thick / 1000));
            else if (material == "ZTR")  //Wieland Translucent Light, Medium... LGS723b
            {
                density = (weight * 1000) / (Math.PI * Math.Pow(diam / 2, 2.0) * thick - (41.878 * 2 / 100.2 * diam));  //Without Ledge Equation
                //density = (weight * 1000) / ((Math.PI * Math.Pow(innerDiam, 2) * thick / 4) + (Math.PI * ledgeThick * (Math.Pow(diam, 2) - Math.Pow(innerDiam, 2)) / 4));
            }
            else if (material == "ZMO")  //Wieland MO LGS723d
            {
                density = (weight * 1000) / (Math.PI * Math.Pow((diam / 2), 2.0) * thick - (0.4 * Math.PI * diam));     //Without Ledge Equation
                //density = (weight * 1000 * .9975) / ((Math.PI * Math.Pow(innerDiam, 2) * thick / 4) + (Math.PI * ledgeThick * (Math.Pow(diam, 2) - Math.Pow(innerDiam, 2)) / 4));
            }
            else if (material == "ZT")  //Wieland T0,T1...  LGS723e
            {
                density = weight * 0.9955 / ((Math.PI * Math.Pow(diam, 2.0) * thick / 4 / 1000 - 2 * 2 * Math.PI * (diam / 2 / 1000) * (0.27 / 2)));  //Without Ledge Equation
                //density = (weight * 1000 * .9955) / ((Math.PI * Math.Pow(innerDiam, 2) * thick / 4) + (Math.PI * ledgeThick * (Math.Pow(diam, 2) - Math.Pow(innerDiam, 2)) / 4));
            }
            else if (material == "ZTG")  //Wieland Green Light, Medium... LGS723c
                density = (weight * 1000) / ((Math.PI * (Math.Pow((diam / 2), 2.0) * thick) - (41.878 * 2 / 100.2 * diam)));
            else if (material == "ZG")  //Wieland Green T0,T1,T2... LGS723f
                density = weight / ((Math.PI * Math.Pow(diam, 2) * thick / 4 / 1000 - 2 * 2 * Math.PI * (diam / 2 / 1000) * (0.27 / 2)));
            else if (material == "ZCLT")
                density = weight * 0.9955 / ((Math.PI * Math.Pow(diam, 2.0) * thick / 4 / 1000 - 2 * 2 * Math.PI * (diam / 2 / 1000) * (0.27 / 2)));  //Without Ledge Equation
            else if (material == "ZCMO")
                density = (weight * 1000) / (Math.PI * Math.Pow((diam / 2), 2.0) * thick - (0.4 * Math.PI * diam));     //Without Ledge Equation
            else if (material == "ZMU")
                density = (weight * 1000) / (Math.PI * Math.Pow((diam / 2), 2.0) * thick);
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
            bool leave = false;
            while (leave==false)
            {
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
                    leave = true;
                }
                catch
                {
                    Console.WriteLine("G:/LabX_Export/LabsQ_LabX_Integration/WielandID Database.csv");
                    Console.WriteLine("in use or incorrectly formatted");
                    Console.WriteLine("Press any key to try again or q to quit...\n");
                    ConsoleKeyInfo keyStroke = Console.ReadKey(false);
                    if (keyStroke.KeyChar == 'q')
                        Environment.Exit(0);
                }
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
        static String GenerateRFIDOutput(int discnum, string batchID, string theoThick, int matNum, string material, string theoDiam, double EF, int simultaneousPrints, int batchSize, string RFIDLocation, string shade, string shadeAlias,double measDiam, string RFIDHeader, double shrinkage,bool retain, bool printToggle)
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
                finalExport = String.Concat(Environment.NewLine, discID, "\t", efDigits, "\t", "98.5", "\t", theoThick, "\t", matNum, "\t", barcodeNutzdaten, "\t", WielandID, "\t", stringNum, "\t\t", batchID, "\t", Math.Round(shrinkage, 3), "\tH\t\t", batchID, "\t\t4418753237\t", barcode, "\t", encodedCode, "\t", shadeAlias, "\t"+shadeAlias2+"\t");
            else 
                finalExport = String.Concat(Environment.NewLine, discID, "\t", efDigits, "\t", heightAlias, "\t", theoThick, ".00\t", matNum, "\t", barcodeNutzdaten, "\t", batchID, "\t", stringNum, "\t\t", batchID, "\t", Math.Round(measDiam, 2), "\tH\t\t", batchID, "\t\t4418753237\t", barcode, "\t", encodedCode, "\t", shadeAlias, "\t\t");

            File.AppendAllText(RFIDLocation + batchID + "_rfid_complete.csv", finalExport);

            //export rfid to smaller sheet
            ProcessStartInfo theDruck = new ProcessStartInfo();
            theDruck.CreateNoWindow = true;
            theDruck.UseShellExecute = false;
            theDruck.WindowStyle = ProcessWindowStyle.Hidden;
            if(simultaneousPrints==1)
            {
                File.WriteAllText(RFIDLocation + batchID + "_rfid_" + discnum + ".csv", RFIDHeader);
                File.AppendAllText(RFIDLocation + batchID + "_rfid_" + discnum + ".csv", finalExport);
                
                {
                    if (printToggle)
                    {
                        try
                        {
                            //put in command to start rfid label printer
                            if (material == "ZFC") //Zirlux
                            {

                                theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";
                                if (retain)
                                {
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Zirlux FC2 With Ring_Retain.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                                    Console.WriteLine(thingOuttaSpec);
                                    Console.WriteLine("Retain Label is Now Being Printed");
                                }
                                else
                                {
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                                    Console.WriteLine("C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe" + " /P \"G:\\Topex_Printer\\Zirlux Label Templates\\Zirlux FC2\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv \"  /start /hidden");
                                    Console.WriteLine("\n\nDisc is in spec\nLabel is printing");
                                }
                            }

                            //new wieland ZMU, ZCLT,ZCMO
                            else if (material == "ZMU" || material == "ZCLT" || material == "ZCMO")
                            {
                                theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";
                                if (retain)
                                {
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Ivoclar_CE0123.txt_Retain.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                                    Console.WriteLine(thingOuttaSpec);
                                    Console.WriteLine("Retain Label is Now Being Printed");
                                }
                                else
                                {
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Ivoclar_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                                    Console.WriteLine("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Topex_Printer\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv \"  /start /hidden");
                                    Console.WriteLine("\n\nDisc is in spec\nLabel is printing");
                                }
                            }

                            else  //Old Wieland
                            {
                                theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";
                                if (retain)
                                {
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Wieland_CE0123_Retain.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                                    Console.WriteLine(thingOuttaSpec);
                                    Console.WriteLine("Retain Label is Now Being Printed");
                                }
                                else
                                {
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv\"  /start /hidden";
                                    Console.WriteLine("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Topex_Printer\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + discnum + ".csv \"  /start /hidden");
                                    Console.WriteLine("\n\nDisc is in spec\nLabel is printing");
                                }
                            }
                            using (Process exeProcess = Process.Start(theDruck))
                            {
                                exeProcess.WaitForExit(20000);
                            }
                            if(discnum==1||discnum==batchSize)
                            {
                                using (Process exeProcess = Process.Start(theDruck))
                                {
                                    exeProcess.WaitForExit(10);
                                }
                            }

                        }
                        catch
                        {
                            Console.WriteLine("Failed to find DruckerStation.exe");
                            Environment.Exit(0);
                        }
                    }
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
                    if (printToggle)
                    {
                        try
                        {
                            if (material == "ZFC") //Zirlux
                            {
                                theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";
                                if (retain)
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Zirlux FC2 With Ring_Retain.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv\"  /start /hidden";
                                else
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv\"  /start /hidden";
                                Console.WriteLine("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Zirlux FC2 With Ring.txt\" /RFID Off  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv \"  /start /hidden");
                                //Console.ReadKey();
                            }

                            //new wieland ZMU, ZCLT,ZCMO
                            else if (material == "ZMU" || material == "ZCLT" || material == "ZCMO")
                            {
                                theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";
                                if (retain)
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Ivoclar_CE0123.txt_Retain.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv\"  /start /hidden";
                                else
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Ivoclar_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv\"  /start /hidden";
                                Console.WriteLine("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Ivoclar_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv \"  /start /hidden");
                            }

                            else  //Old Wieland ZMO,ZT,ZTR
                            {
                                theDruck.FileName = @"C:\Program Files (x86)\Wieland RFID PrinterStation\DruckerStation.exe";
                                if (retain)
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Wieland_CE0123_Retain.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv\"  /start /hidden";
                                else
                                    theDruck.Arguments = " /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv\"  /start /hidden";
                                Console.WriteLine("\"C:\\Program Files (x86)\\Wieland RFID PrinterStation\\DruckerStation.exe\" /P \"G:\\Prod_Labels\\Zirconia\\RFID Labels\\TOPEX_BITMAPS\\Wieland_CE0123.txt\" /RFID On  /B \"" + RFIDLocation + batchID + "_rfid_" + Convert.ToString(lowNum) + "-" + Convert.ToString(highNum) + ".csv \"  /start /hidden");
                                //Console.ReadKey();
                            }
                            using (Process exeProcess = Process.Start(theDruck))
                            {
                                exeProcess.WaitForExit();
                            }
                            if (discnum == 1 || discnum == batchSize)
                            {
                                using (Process exeProcess = Process.Start(theDruck))
                                {
                                    exeProcess.WaitForExit(10);
                                }
                            }
                        }
                        catch
                        {
                            Console.WriteLine("Failed to find DruckerStation.exe");
                            Environment.Exit(0);
                        }
                    }

                    //put in command to start rfid label printer
                    
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
        
        static bool SpecCheck(string material, int matNum, double measDiam, double measThick, double PSDensity, double innerDiam, double ledgeThick, double ledgeOffset, double concentricity, double EF, string materialID)
        {
            //checks to see if all the data for the disc is in spec
            //returns location of template file to use 
            //out of spec returns retain template in spec returns normal label
            bool isOutOfSpec = false;
            string specListLocation = "G:/LabX_Export/LabsQ_LabX_Integration/Spec_List.csv";
            var column1 = new List<string>();
            var column2 = new List<string>();
            var column3 = new List<string>();
            var column4 = new List<string>();
            var column5 = new List<string>();
            var column6 = new List<string>();
            var column7 = new List<string>();
            var column8 = new List<string>();
            var column9 = new List<string>();
            var column10 = new List<string>();
            var column11 = new List<string>();
            var column12 = new List<string>();
            var column13 = new List<string>();
            var column14 = new List<string>();
            var column15 = new List<string>();
            var column16 = new List<string>();
            bool leave = false;
            while (leave == false)
            {
                try
                {
                    using (var rd = new StreamReader(specListLocation))
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
                            column10.Add(splits[9]);
                            column11.Add(splits[10]);
                            column12.Add(splits[11]);
                            column13.Add(splits[12]);
                            column14.Add(splits[13]);
                            column15.Add(splits[14]);
                            column16.Add(splits[15]);
                        }
                        leave = true;
                    }
                }
                catch
                {
                    Console.WriteLine(specListLocation);
                    Console.WriteLine("in use or incorrectly formatted");
                    Console.WriteLine("Press any key to try again or q to quit...\n");
                    ConsoleKeyInfo keyStroke = Console.ReadKey(false);
                    if (keyStroke.KeyChar == 'q')
                        Environment.Exit(0);
                }
            }
            int rownum = 0;
            for (int i = 0; i < column1.Count;i++)
            {
                if (column1[i] == materialID)
                    rownum = i;
            }
            //check if thickness is in spec
            if (Math.Round(measThick, Convert.ToInt32(column3[1])) < Math.Round(Convert.ToDouble(column3[rownum]),Convert.ToInt32(column3[1])) || Math.Round(measThick, Convert.ToInt32(column4[1])) > Math.Round(Convert.ToDouble(column4[rownum]),Convert.ToInt32(column4[1])))
            {
                thingOuttaSpec = "Thickness: " + Math.Round(measThick, Convert.ToInt32(column3[1])) + "\n Spec: " + Math.Round(Convert.ToDouble(column3[rownum]), Convert.ToInt32(column3[1])) + "-" + Math.Round(Convert.ToDouble(column4[rownum]), Convert.ToInt32(column4[1]));
                whatsOuttaSpec = whatsOuttaSpec + "Thickness ";
                isOutOfSpec = true;
            }
            //check if diameter is in spec, if zirlux override rounding to 2
         
            if(material=="ZFC")
            {
                if (Math.Round(measDiam, 2) < Math.Round(Convert.ToDouble(column5[rownum]), 2) || Math.Round(measDiam, 2) > Math.Round(Convert.ToDouble(column6[rownum]), 2))
                {
                    thingOuttaSpec = "Diameter: " + Math.Round(measDiam, 2) + "\n Spec: " + Math.Round(Convert.ToDouble(column5[rownum]), 2) + "-" + Math.Round(Convert.ToDouble(column6[rownum]), 2);
                    whatsOuttaSpec = whatsOuttaSpec + "Diameter ";
                    isOutOfSpec = true;
                }
            }
            else
            {
                if (Math.Round(measDiam, Convert.ToInt32(column5[1])) < Math.Round(Convert.ToDouble(column5[rownum]), Convert.ToInt32(column5[1])) || Math.Round(measDiam, Convert.ToInt32(column6[1])) > Math.Round(Convert.ToDouble(column6[rownum]), Convert.ToInt32(column6[1])))
                {
                    thingOuttaSpec = "Diameter: " + Math.Round(measDiam, Convert.ToInt32(column5[1])) + "\n Spec: " + Math.Round(Convert.ToDouble(column5[rownum]), Convert.ToInt32(column5[1])) + "-" + Math.Round(Convert.ToDouble(column6[rownum]), Convert.ToInt32(column6[1]));
                    whatsOuttaSpec = whatsOuttaSpec + "Diameter ";
                    isOutOfSpec = true;
                }
            }
            

            if (material != "ZFC")
            {
                //check if density is in spec, if zenostar multi override rounding to 3
                if(material=="ZMU")
                {
                    if (Math.Round(PSDensity, 3) < Math.Round(Convert.ToDouble(column7[rownum]), 3) || Math.Round(PSDensity, 3) > Math.Round(Convert.ToDouble(column8[rownum]), 3))
                    {
                        thingOuttaSpec = "Pre-Sintered Density: " + Math.Round(PSDensity, Convert.ToInt32(column7[1])) + "\n Spec: " + Math.Round(Convert.ToDouble(column7[rownum]), Convert.ToInt32(column7[1])) + "-" + Math.Round(Convert.ToDouble(column8[rownum]), Convert.ToInt32(column8[1]));
                        whatsOuttaSpec = whatsOuttaSpec + "Density ";
                        isOutOfSpec = true;
                    }
                }
                else
                {
                    if (Math.Round(PSDensity, Convert.ToInt32(column7[1])) < Math.Round(Convert.ToDouble(column7[rownum]), Convert.ToInt32(column7[1])) || Math.Round(PSDensity, Convert.ToInt32(column8[1])) > Math.Round(Convert.ToDouble(column8[rownum]), Convert.ToInt32(column8[1])))
                    {
                        thingOuttaSpec = "Pre-Sintered Density: " + Math.Round(PSDensity, Convert.ToInt32(column7[1])) + "\n Spec: " + Math.Round(Convert.ToDouble(column7[rownum]), Convert.ToInt32(column7[1])) + "-" + Math.Round(Convert.ToDouble(column8[rownum]), Convert.ToInt32(column8[1]));
                        whatsOuttaSpec = whatsOuttaSpec + "Density ";
                        isOutOfSpec = true;
                    }
                }
                
                
                    
                //only check ledge stuff on non-green wieland discs
                /*  OBSOLETE / NOT CHECKED ANYMORE
                if(material=="ZT"||material=="ZTR"||material=="ZMO")
                {
                    //Inner Diameter spec
                    if (Math.Round(innerDiam, Convert.ToInt32(column11[1])) < Math.Round(Convert.ToDouble(column11[rownum]), Convert.ToInt32(column11[1])) || Math.Round(innerDiam, Convert.ToInt32(column12[1])) > Math.Round(Convert.ToDouble(column12[rownum]), Convert.ToInt32(column12[1])))
                    {
                        thingOuttaSpec = "Inner Diameter: " + Math.Round(innerDiam, Convert.ToInt32(column11[1])) + "\n Spec: " + Math.Round(Convert.ToDouble(column11[rownum]), Convert.ToInt32(column11[1])) + "-" + Math.Round(Convert.ToDouble(column12[rownum]), Convert.ToInt32(column12[1])) ;
                        return true;
                    }
                    //Ledge Thickness spec
                    if (Math.Round(ledgeThick, Convert.ToInt32(column13[1])) < Math.Round(Convert.ToDouble(column13[rownum]), Convert.ToInt32(column13[1])) || Math.Round(ledgeThick, Convert.ToInt32(column14[1])) > Math.Round(Convert.ToDouble(column14[rownum]), Convert.ToInt32(column14[1])))
                    {
                        thingOuttaSpec = "Ledge Thickness: " + Math.Round(ledgeThick, Convert.ToInt32(column13[1])) + "\n Spec: " + Math.Round(Convert.ToDouble(column13[rownum]), Convert.ToInt32(column13[1])) + "-" + Math.Round(Convert.ToDouble(column14[rownum]), Convert.ToInt32(column14[1]));
                        return true;
                    }
                    //Center Offset Max
                    if (Math.Round(concentricity, Convert.ToInt32(column15[1])) > Math.Round(Convert.ToDouble(column15[rownum]), Convert.ToInt32(column15[1])))
                    {
                        thingOuttaSpec = "Concentricity: " + Math.Round(concentricity, Convert.ToInt32(column15[1])) + "\n Spec: <" + Math.Round(Convert.ToDouble(column15[rownum]), Convert.ToInt32(column15[1]));
                        return true;
                    }
                    //Vertical Offset Max
                    if (Math.Round(ledgeOffset, Convert.ToInt32(column16[1])) > Math.Round(Convert.ToDouble(column16[rownum]), Convert.ToInt32(column16[1])))
                    {
                        thingOuttaSpec = "Ledge Offset: " + Math.Round(ledgeOffset, Convert.ToInt32(column16[1])) + "\n Spec: <" + Math.Round(Convert.ToDouble(column16[rownum]), Convert.ToInt32(column16[1]));
                        return true;
                    }
                    
                }
                */
            }
            //check for EF in zirlux discs
            else
                if (Math.Round(EF, Convert.ToInt32(column9[1])) < Math.Round(Convert.ToDouble(column9[rownum]), Convert.ToInt32(column9[1])) || Math.Round(EF, Convert.ToInt32(column10[1])) > Math.Round(Convert.ToDouble(column10[rownum]), Convert.ToInt32(column10[1])))
                {
                    thingOuttaSpec = "EF: " + Math.Round(EF, Convert.ToInt32(column9[1])) + "\n Spec: " + Math.Round(Convert.ToDouble(column9[rownum]), Convert.ToInt32(column9[1])) + "-" + Math.Round(Convert.ToDouble(column10[rownum]), Convert.ToInt32(column10[1]));
                    whatsOuttaSpec = whatsOuttaSpec + "EF ";
                    isOutOfSpec = true;
                }


            if (isOutOfSpec) isConforming = "Nonconform";
 
            return isOutOfSpec;
        }
        static void PrintBoxLabels(int discnum,string theoThick,int matNum,string batchID,string material,string materialID,int batchsize,string templateLocation,int numLabels, string shade, string fullMaterialID, string packagedMatNum)
        {
            var cmdFile = new StringBuilder();
            var cmdFile2 = new StringBuilder();

            string shadeCopy = shade;
            if (Char.IsDigit(shade[1])) shade = shade.Insert(1, " ");
            else if (Char.IsDigit(shade[2])) shade = shade.Insert(2, " ");
            if (shade == "Sun-Chroma") shade = "T sc";
            if (shade == "Sun") shade = "T s";
            
            if(material=="ZMO"||material=="ZT")
                cmdFile.AppendLine("LABELNAME = \"" + templateLocation + "Zenostar Box Label Template.lab\"");
            else if(material=="ZTR")
                cmdFile.AppendLine("LABELNAME = \"" + templateLocation + "Zenostar Trans Box Label Template.lab\"");
            else if (material=="ZMU")
                cmdFile.AppendLine("LABELNAME = \"" + templateLocation + "Zircad Disc Template.lab\"");
            else if (material == "ZCLT")
                cmdFile.AppendLine("LABELNAME = \"" + templateLocation + "Zircad Disc Template.lab\"");
            else if (material == "ZCMO")
                cmdFile.AppendLine("LABELNAME = \"" + templateLocation + "Zircad Disc Template.lab\"");
            else 
                cmdFile.AppendLine("LABELNAME = \"" + templateLocation + "Zirlux Box Label Template.lab\"");

            cmdFile.AppendLine("PRINTER= \"Copy of Datamax-O'Neil I-4606e Mark II,USB004\"");
            cmdFile.AppendLine("Thickness = \"" + theoThick + "\"");
            cmdFile.AppendLine("Material Number = \"" + packagedMatNum + "\"");
            cmdFile.AppendLine("Lot Number = \"" + batchID + "\"");

            if (material != "ZFC") 
                cmdFile.AppendLine("Material Description = \"" + shade + "\"");
            else
                cmdFile.AppendLine("Material Description = \"" + fullMaterialID.Replace('_', ' ') + "\"");

            cmdFile.AppendLine("LABELQUANTITY = \"" + numLabels + "\"");

            File.WriteAllText(templateLocation+"Label Output\\Box Output_"+discnum+".cmd",cmdFile.ToString());

            if (material == "ZT" || material == "ZMO") 
            {
                cmdFile2.AppendLine("LABELNAME = \"" + templateLocation + "Zenostar Barcode Label Template.lab\"");
                cmdFile2.AppendLine("PRINTER = \"Datamax-O'Neil I-4606e Mark II,USB003\"");
                cmdFile2.AppendLine("Material Number = \"" + packagedMatNum + "\"");
                cmdFile2.AppendLine("Lot Number = \"" + batchID + "\"");
                cmdFile2.AppendLine("Piece Count = \"1 pc.\"");
                cmdFile2.AppendLine("Barcode Label Identifier = \"" + fullMaterialID.Replace('_', ' ') + "\"");
                cmdFile2.AppendLine("LABELQUANTITY = \"" + numLabels + "\"");
                File.WriteAllText(templateLocation + "Label Output\\Barcode Output_" + discnum + ".cmd", cmdFile2.ToString());
            }
            try
            {
                if (discnum == 1) 
                {
                    File.WriteAllText("G:/Equipment/GiS-Topex RFID Printer/Generated RFID Files/" + shadeCopy + "/" + batchID + "/Box Output_" + batchID + ".cmd", cmdFile.ToString());
                    if(material!="ZFC")
                        File.WriteAllText("G:/Equipment/GiS-Topex RFID Printer/Generated RFID Files/" + shadeCopy + "/" + batchID + "/Barcode Output_" + batchID + ".cmd", cmdFile2.ToString());

                    ProcessStartInfo csStart = new ProcessStartInfo();
                    csStart.CreateNoWindow = true;
                    csStart.UseShellExecute = false;
                    csStart.WindowStyle = ProcessWindowStyle.Hidden;
                    csStart.FileName = @"C:\Program Files (x86)\Teklynx\CODESOFT 2014\CS.exe";
                    csStart.Arguments = " /CMD " + templateLocation + "Label Output";
                    using (Process exeProcess = Process.Start(csStart))
                        exeProcess.WaitForExit(10000);
                }
            }
            catch
            {
                Console.WriteLine("Opening CS.exe Failed");
                Console.WriteLine("Program will now end\nPress any key");
                Console.ReadKey();
            }

            int hWnd;
            Process[] processRunning = Process.GetProcesses();
            foreach (Process pr in processRunning)
            {
                if (pr.ProcessName == "CS")
                {
                    hWnd = pr.MainWindowHandle.ToInt32();
                    ShowWindow(hWnd, 6);
                }
                if (pr.ProcessName == "Lppa")
                {
                    hWnd = pr.MainWindowHandle.ToInt32();
                    ShowWindow(hWnd, 6);  //6 is the number to minimize
                }
            }

            if(discnum==batchsize)
            {
                Console.WriteLine("\n\nClosing Label Printing Software...");
                Thread.Sleep(15000);

                foreach (var process in Process.GetProcessesByName("Lppa"))
                    process.Kill();

                foreach (var process in Process.GetProcessesByName("CS"))
                    process.Kill();
            }
        }

        [DllImport("User32")]
        private static extern int ShowWindow(int hwnd, int nCmdShow);
        

    }

}
