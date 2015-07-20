using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace PerfParse {
  class Program {
    static void Main(string[] args) {
      string nmon = "";
      string xlFile = "";
      string sType = "NF";
      string sStatus = "NF";
      string cID = "0";
      string here = Directory.GetCurrentDirectory();
      bool found = false;
      if (args.Length >= 5) {
        //dowork
        nmon = @here + @"\" + args[0];
        xlFile = @here+@"\"+args[1];
        cID = args[2];
        sType = args[3];
        sStatus = args[4];
        found = true;
      } else {
        Console.WriteLine("Please enter the appropriate arguments."+Environment.NewLine+"PerfParse.exe  <rawdata.[nmon|zip(blg)]> <perf_collection.xlsx> <compile ID#> <server(DAC,APP,etc)> <status(Online,Backup)>");
        Console.ReadKey();
      }

      if (found) {
        if (nmon.IndexOf("nmon") > 0) {
          List<process> processes = readSherpa(nmon, false); //get nmon processes
          List<string> omittedP = updateXL(processes, xlFile, cID, sType, sStatus);
          if (omittedP.Count > 0) {
            string omist = @here + @"\omissions.txt";
            using (StreamWriter sw = File.CreateText(omist)) {
              sw.WriteLine("The following Commands were not found in the spreadsheet and omitted:");
              foreach (string ot in omittedP) {
                sw.WriteLine(ot);
              }
            }
            //print omitted processes to text file omissions.txt
          } else {
            Console.WriteLine("Check PerfParse arguments for spelling, had troubles finding specified location in Excel.");
          }
        } else {//ZIP with BLGs
          string zipDir = "";
          string DirName = "";
          bool singles = false;
          try {
            using (ZipArchive archive = ZipFile.OpenRead(nmon)) {
              foreach (ZipArchiveEntry entry in archive.Entries) {
                string exName = "";
                if (entry.Name.Contains(".blg")) {
                  try {//can you unzip that file?
                    //OVERWRITE SAME FOLDER MUHGAWD pls fix
                    string fzName = entry.FullName;
                    int ixf = fzName.IndexOf("/");
                    if (ixf >= 0) {
                      exName = fzName.Substring(0, ixf);
                      DirName = exName;
                      zipDir = Path.Combine(nmon.Substring(0, nmon.LastIndexOf('\\')), exName);
                    } else {
                      singles = true;
                      DirName = "";
                      exName = fzName;
                      zipDir = nmon.Substring(0, nmon.LastIndexOf('\\'));
                    }
                    // string nfhx = Path.Combine(nmon.Substring(0, nmon.LastIndexOf('\\')), exName);
                     
                     string nfhx = Path.Combine(nmon.Substring(0, nmon.LastIndexOf('\\')), exName);
                    //int inc = 1;
                    while (Directory.Exists(nfhx)) {
                      nfhx += "x";
                    }
                    if (ixf >= 0) {
                      archive.ExtractToDirectory(nfhx);
                    } else {
                      entry.ExtractToFile(nfhx);
                    }
                    //
                    
                    Console.WriteLine("Unzipped " + exName);
                    break;
                  } catch {
                    //errors.Add("Had trouble unzipping " + exName +" please check for completion.");
                  }
                }
              }
            }

            string[] blgfiles = Directory.GetFiles(zipDir, "*.blg", SearchOption.AllDirectories);
            string allBLGs = string.Join(" ", blgfiles);
            string filterText = @here + @"\CB745.txt";
            using (StreamWriter sw = File.CreateText(filterText)) {
              sw.WriteLine(@"\Process(*)\% Processor Time");
            }
            Console.WriteLine("Relogging Windows PerfMons.");
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "CMD.exe";
            startInfo.Arguments = "/C relog " + allBLGs + " -o CombinedLog.blg";
            //csvfiles.Add(blgfile.Substring(0, blgfile.Length - 4) + ".csv");
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();
            if (singles) {
              File.Delete(allBLGs);
            } else {
              Directory.Delete(zipDir, true);
            }
            startInfo.Arguments = "/C relog CombinedLog.blg -cf " + filterText + " -f csv -o " + DirName + "CPU.csv"; ;
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();
            File.Delete(filterText);
            using (StreamWriter sw = File.CreateText(filterText)) {
              sw.WriteLine(@"\Process(*)\Working Set");
            }
            startInfo.Arguments = "/C relog CombinedLog.blg -cf " + filterText + " -f csv -o " + DirName + "MEM.csv"; ;
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();
            File.Delete(filterText);
            //%DirName%CPU.csv
            //%DirName%MEM.csv
            List<process> processes = getBLGdata(@here + "\\" + DirName + "CPU.csv", @here + "\\" + DirName + "MEM.csv");
            double cores = 24;
            foreach (process x in processes) {
              if (x.command == "_Total") {
                cores = x.cpuPercent.Max();
                break;
              }
            }
            for (int i = 0; i < processes.Count(); i++) {
              if (processes[i].command != "_Total") {
                for (int j = 0; j < processes[i].cpuPercent.Count; j++) {
                  processes[i].cpuPercent[j] = processes[i].cpuPercent[j] / cores;
                }
              }
            }
            List<string> omittedP = updateXL(processes, xlFile, cID, sType, sStatus);
            if (omittedP.Count > 0) {
              string omist = @here + @"\omissions.txt";
              using (StreamWriter sw = File.CreateText(omist)) {
                sw.WriteLine("The following Commands were not found in the spreadsheet and omitted:");
                foreach (string ot in omittedP) {
                  sw.WriteLine(ot);
                }
              }
            } else {
              Console.WriteLine("Check PerfParse arguments for spelling, had troubles finding specified location in Excel.");
            }
            File.Delete(DirName + "CPU.csv");
            File.Delete(DirName + "MEM.csv");
            if (args.Length == 5)
              File.Delete("CombinedLog.blg");
          } catch {
            Console.WriteLine("Woah - Your System is likely not updated to support .NET Framework 4.5.");
          }
        }
        
      }
    }

    private static List<process> getBLGdata(string p1, string p2) {
      List<string[]> mems = getRows(p2);
      List<string[]> cpus = getRows(p1);
      List<process> processes = new List<process>();
      int cols = cpus[0].Length;
      int rows = cpus.Count();
      for (int i = 1; i < cols; i++) {//start on column 1
        for (int j = 1; j < rows; j++) {//start on row 2
          string cmd = mems[0][i];
          int fM = cmd.IndexOf("(");
          int lM = cmd.IndexOf(")");
          string m = mems[j][i];
          m = (Convert.ToDouble(m) / 1024).ToString();
          string c = cpus[j][i];
          string cm = cmd.Substring(fM + 1, lM - fM - 1);
          if(cm.IndexOf("#")>=0){
            cm = cm.Remove(cm.IndexOf("#"));
          }
          bool found = false;
          foreach (process x in processes) {
            if (x.command == cm) { 
            //if (x.pid == i) {
              x.add(m, c);
              found = true;
              break;
            }
          }
          if (!found) {            
            process item = new process(i.ToString(), m, c, cm);
            processes.Add(item);
          }          
        }
      }
        return processes;
    }

    private static List<string[]> getRows(string file) {
      List<string[]> datas = new List<string[]>();
      char[] delimiters = new char[] { ',' };
      using (StreamReader reader = new StreamReader(file)) {
        int ln = 0;
        while (true) {
          string line = reader.ReadLine();
          if (line == null) {
            break;
          }
          line = line.Replace("\"", "");
          if (ln != 1) {
            string[] row = line.Split(delimiters);
            string[] data = row.Select(x => x.Replace(" ", "0")).ToArray();
            datas.Add(data);
          }
          ln++;
        }
      }
      return datas;
    }

    static List<string> updateXL(List<process> processes, string fileB, string cID, string type, string status) {
      List<string> omissions = new List<string>();
      Excel.Application excelApp = null;
        //Excel.Workbook OldBook = null;
        Excel.Workbook NewBook = null;
        //Excel.Worksheet OldSheet = null;
        Excel.Worksheet NewSheet = null;
        //Excel.Worksheet dtSheet = null;
        //Excel.Range R1 = null;
        //Excel.Range R2 = null;
        try {
          excelApp = new Excel.Application(); ;
          excelApp.DisplayAlerts = false;
          //OldBook = excelApp.Workbooks.Open(fileA, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          NewBook = excelApp.Workbooks.Open(fileB, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

          NewSheet = NewBook.Sheets[1];
          int NewMR = NewSheet.UsedRange.Rows.Count;
          int NewMC = NewSheet.UsedRange.Columns.Count;
          object[,] NewData = NewSheet.get_Range("A4:C" + NewMR).Value;//get full new WR list
          List<string> newP = getColumn(NewData, 1);
          List<string> newT = getColumn(NewData, 2);
          List<string> newS = getColumn(NewData, 3);
          List<string> Nsheet = new List<string>();
          for (int i = 0; i < (NewMR-3); i++) {
            Nsheet.Add(newP[i] + "," + newT[i] + "," + newS[i]);
          }

          int columnF = 4;
          for (int i = columnF; i <= NewMC; i++) {
            double lft = 0;
            try { lft = (double)(NewSheet.Cells[1, i] as Excel.Range).Value; } catch (Exception e) { string m = e.Message; }
            if ( lft == Convert.ToDouble(cID)) {//found Compile ID column
              foreach (process x in processes) {//foreach recorded nmon process
                int index = getSIndex(Nsheet, x, type, status);//find the row that the process exists on
                if (index >= 0) {//if the process exists
                  //column i
                  //row index(0)
                  //process x
                  NewSheet.Cells[index + 4, i] = x.cpuPercent.Average();
                  NewSheet.Cells[index + 4, i + 1] = (x.resset.Max()/1024).ToString();
                } else {
                  omissions.Add(x.command);
                }
              }
            }
          }
          if (processes.Count == omissions.Count) {
            omissions = new List<string>();
          }
          NewBook.SaveAs(Directory.GetCurrentDirectory()+"\\outputPerf", XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
          NewBook.Close(true, Type.Missing, Type.Missing);
         // OldBook.Close(false, Type.Missing, Type.Missing);
          excelApp.Quit();
        } catch (Exception e){
          Console.WriteLine("Danger to manifold! Check for format issues.");
          Console.WriteLine(e.Message);
          Console.ReadKey();
        } finally {
          releaseObject(NewSheet);
          //releaseObject(OldSheet);
          releaseObject(NewBook);
          //releaseObject(OldBook);
          releaseObject(excelApp);
        }


      return omissions;
    }
    static int getSIndex(List<string> values1, process x, string type, string status) {//Nsheet, x, type, status
      int index = -1;
      for (int i = 0; i < values1.Count; i++) {
        //if (values1[i] == value1 && values2[i] == value2) {
        if(values1[i]==(x.command+","+type+","+status)){
          index = i;
          break;
        }
      }
      return index;
    }
    static List<string> getColumn(object[,] dataTable, int column) {
      List<string> data = new List<string>();
      int maxR = dataTable.GetLength(0);
      for (int i = 1; i <= maxR; i++) {
        data.Add(dataTable[i, column].ToString());
      }
      return data;
    }
    static void releaseObject(object obj) {
      try {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      } catch {
        obj = null;
      } finally {
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
    }

    static List<process> readSherpa(string filename, bool print) {
      string line;
      string host = "new";
      int cores = 0;
      List<process> processes = new List<process>();
      List<string> summary = new List<string>();
      List<string> times = new List<string>();
      List<double> cpu_all_usr = new List<double>();
      List<double> memsum = new List<double>();
      List<string> disklabels = new List<string>();
      List<double> disksizes = new List<double>();
      List<string> disksizesb = new List<string>();
      List<string> netwuts = new List<string>();
      List<List<string>> topList = new List<List<string>>();
      List<double>[] diskbusy;
      List<double>[] netties;
      string[] dix;

      string datime = "";
      string ddate = "";
      string ttime = "";
      //string warnings = "";

      using (StreamReader reader = new StreamReader(filename)) {

        /*read in each line of text then do stuff with it*/
        //small while loop only does maybe 50lines before breaking
        while ((line = reader.ReadLine()) != null) {//this is the prelim loop to make the primary loop go quicker
          summary.Add(line);
          string[] values = line.Split(',');
          if (values[1] == "time") {
            if (values[0] == "AAA")
              ttime = values[2];
            datime = String.Join("", values[2].Split(new[] { ':', '.' }));
          }
          if (values[1] == "date") {
            if (values[0] == "AAA")
              ddate = values[2];
            datime = String.Join("", values[2].Split('-')) + "_" + datime;
          }
          if (values[1] == "host")
            host = values[2];
          if (values[1] == "cpus")
            cores = Convert.ToInt32(values[2]);
          if (values[0] == "NET") {//first line of NET data from the file
            foreach (string nets in values.Skip(2)) { //for all the nets presented on this line (skipping the first 2 garbage lines)
              if (nets != "") netwuts.Add(nets);//all the things, each iface, each bond, eths, los..  everything from the ifconfig
            }
          }
          if (values[0] == "DISKBUSY") {//first line of DISKBUSY holds disk names
            foreach (string diskN in values.Skip(2)) { //for all the disk labels presented on this line (skipping the first 2 garbage lines)
              if (diskN != "") disklabels.Add(diskN);//all sd and dm partitions, just keep it all in there
            }
          }
          if (values[0] == "BBBP") {
            if (values[2] == "/proc/partitions") {
              try {
                dix = values[3].Split(new[] { ' ', '\"' }, StringSplitOptions.RemoveEmptyEntries);
                if (dix[0] != "major") {
                  disksizes.Add(Convert.ToDouble(dix[2]) / 1000);
                  disksizesb.Add(dix[3]);
                }
              } catch { }
            } else if (values[2] == "/proc/1/stat")
              break;
          }
        }//some background info was gathered from AAA


        netties = new List<double>[netwuts.Count()];
        for (int i = 0; i < netties.Count(); i++) {
          netties[i] = new List<double>();//so many I dont even
        }//we now have netwuts.count netties[]s; each netties is a double list we can add each(EVERY SINGLE) line nmon records

        diskbusy = new List<double>[disklabels.Count()];
        for (int i = 0; i < disklabels.Count(); i++) {
          diskbusy[i] = new List<double>();//almost as many I dont even
        }//we now have disklabels.count diskbusy[]s; each diskbusy is a double list we can add each(EVERY SINGLE) line nmon records
        List<UARG> uargs = new List<UARG>();
        while ((line = reader.ReadLine()) != null) { //Got all the prelim done, now do the rest of the file
          string[] values = line.Split(',');
          /*switch was faster than an if block*/
          try {
            switch (values[0]) {
              case "ZZZZ":
                times.Add(values[2] + " " + values[3]);
                break;
              case "UARG":
                double uargPID = 0;
                string uargCMD = "";
                try { uargPID = Convert.ToDouble(values[2]);
                uargCMD = values[4];
                if (uargCMD.IndexOf(" ") >= 0) {
                  UARG tU = new UARG(uargPID, uargCMD.Substring(0, uargCMD.IndexOf(" ")));
                  uargs.Add(tU);
                } else {
                  UARG tU = new UARG(uargPID, uargCMD);
                  uargs.Add(tU);
                }
                  
                } catch { }

                break;
              case "TOP":
                List<string> topstuff = new List<string>();
                //TOP,+PID,Time,%CPU,%Usr,%Sys,Size,ResSet,ResText,ResData,ShdLib,MajorFault,MinorFault,Command
                //TOP,0031885,T0050,92.9,89.2,3.7,1416768,1105388,144,0,143692,34,0,osii_dbms_adapt  
                topstuff.Add(values[2].Substring(1, values[2].Length - 1));//time in front of topstuff
                for (int i = 1; i < values.Count(); i++) {
                  if (i != 2) {//skip time
                    topstuff.Add(values[i]);//add each value starting from 1 (skipping 2)
                    //Time,+PID,%CPU,%Usr,%Sys,Size,ResSet,ResText,ResData,ShdLib,MajorFault,MinorFault,Command
                    //0050,0031885,92.9,89.2,3.7,1416768,1105388,144,0,143692,34,0,osii_dbms_adapt
                  }
                }
                topList.Add(topstuff);
                string TOPpid = topstuff[1];
                int TPID = Convert.ToInt32(TOPpid);
                string TOPres = topstuff[6];
                string TOPcpu = topstuff[2];
                bool found = false;

                foreach (process x in processes) {
                  if (x.command == topstuff[12]) {//find by command
                    x.add(TOPres, TOPcpu);
                    found = true;
                    break;
                  }
                  
                  //if (x.pid == TPID) {//find by PID
                  //  x.add(TOPres, TOPcpu);
                  //  found = true;
                  //  break;
                  //}
                }
                if (!found) {
                  process t = new process(TOPpid, TOPres, TOPcpu, topstuff[12]);
                  processes.Add(t);
                }
                break;
              case "CPU_ALL":
                if (values[2] != "User%") {
                  cpu_all_usr.Add((Convert.ToDouble(values[2]) + Convert.ToDouble(values[3])));
                }
                break;
              case "MEM":
                if (values[2] != "memtotal") {
                  memsum.Add(100.0 * (1 - ((Convert.ToDouble(values[6]) + Convert.ToDouble(values[11]) + Convert.ToDouble(values[14])) / Convert.ToDouble(values[2]))));

                }
                break;
              case "NET":
                Parallel.ForEach(values.Skip(2), (nets, y, i) => {
                  if (nets != "") netties[i].Add(Convert.ToDouble(nets));
                });
                break;
              case "DISKBUSY":
                Parallel.ForEach(values.Skip(2), (disk, y, i) => {
                  diskbusy[i].Add(Convert.ToDouble(disk));
                });
                break;
              //etc
              default: //poison buckets barf pile
                break;
            }//end switch
          } catch (Exception e) {
            string m = e.Message;
          }
        }//end while
        for (int i = 0; i < processes.Count; i++) {
          if (processes[i].command.Length >= 14) {
            int uI = findPID(Convert.ToDouble(processes[i].pid), uargs);
            if (uI >= 0) {
              processes[i].command = uargs[uI].CMD;
            }
          }
      }

        }

      

        return (processes);

      }//done file handling
    static int findPID(double PID, List<UARG> uargs) {
      int index = -1;
      for (int i = 0; i < uargs.Count; i++) {
        if (uargs[i].PID == PID) {
          index = i;
          break;
        }
      }
        return index;
    }

  }
}
