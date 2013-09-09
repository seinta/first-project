  
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace WindowsFormsApplication15
{
      
    public partial class Form1 : Form
    {

        //GlobalKeyboardHook gHook; 
         
        //Excel.Application xlApp;                       //orizw tis excel metavlites gia to record
        //Excel.Workbook xlWorkBook;
        //Excel.Worksheet xlWorkSheet;
        //Excel.Range range;
        List<employment> list = new List<employment>();   //orizw tin lista ipallilwn mesa sto programma
        String str;                                      //to string sto opoio grafw osa pairnw apo to excel arxeio
        int rCnt = 0;                                   //oi deiktes grammwn tou excel
        int cCnt = 0;                                //oi deiktes stilwn t excel
        //Excel.Application oXL;                      
        //Excel._Workbook oWB;                           //orizw tis excel metavlites gia to report
        //Excel._Worksheet oSheet;
        //Excel.Range oRng;
        //Excel.Application XL;
        //Excel._Workbook WB;
        //Excel._Worksheet Sheet;                        //orizw tis excel metavlites gia to miniaio report;
        //Excel.Range Rng;
        int countexcel=2;
        int row;
        


        
       
        


        class employment
        {
           
            public string id;                                   
            public string namesurname;
          
            public string timein;
            public string timeout;
            public int inside;        //inside % 2 == 0 tote mesa alliws einai eksw
            public int indexoflistbox;         
            public int numberofbreaks;
            public TimeSpan starttime;
            public TimeSpan breaktime;
            public TimeSpan fullbreaktime;
            public TimeSpan fullworktime;
            public TimeSpan lasttime;
            public TimeSpan overtime;
            public double fullworksec;
            public double fullbreaksec;
            public double starttimesec;
            public double breaktimesec;
            public double worktimesec;
            public double lasttimesec;
            public employment()
            {
                id = null;
                namesurname = null;
                
                inside = 0;
                numberofbreaks = 0;
               
            }
        }
        public Form1()
        {
            
            InitializeComponent();
            Excel.Application xlApp;                       //orizw tis excel metavlites gia to record
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            Excel.Application oXL;
            Excel._Workbook oWB;                           //orizw tis excel metavlites gia to report
            Excel._Worksheet oSheet;
            Excel.Range oRng;
            Excel.Application XL;
            Excel._Workbook WB;
            Excel._Worksheet Sheet;                        //orizw tis excel metavlites gia to miniaio report;
            Excel.Range Rng;
            int month;
            string name;
            string monthname;
            DateTime today = DateTime.Today;
            month = today.Month;
            monthname = today.ToString("MMM");
            name = today.ToString("D");
            string monthpath = "c:/Users/dimitris/Dropbox/WhoWorks/" + monthname;
            string fileName = "records.xls";
            string fileName2 = name+" report.xls";
            string sourcePath = "c:/Users/dimitris/Dropbox/WhoWorks";
            string dokimastiko = monthpath+"/"+monthname+" report.xls";
            if (System.IO.Directory.Exists(monthpath)==false)
            {
                System.IO.Directory.CreateDirectory(monthpath);
            }
            if (System.IO.File.Exists(dokimastiko))
            {
                string mpa;
                XL = new Excel.Application();
               // XL.Visible = false;
               // xlWorkBook = xlApp.Workbooks.Open(sourceFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                WB = XL.Workbooks.Open(dokimastiko, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //  oWB.SaveCopyAs("c:/Program Files/WhoWorks/"+name+" report.xls");
                Sheet = (Excel.Worksheet)WB.Worksheets.get_Item(1);
                
                Rng = Sheet.UsedRange;
                int rowcount;
                int exist=0;
                for (rowcount = 1; rowcount <= Rng.Rows.Count; rowcount++)
                {

                    mpa = Convert.ToString((Rng.Cells[rowcount, 12] as Excel.Range).Value2);
                    if (mpa == DateTime.Today.ToString())
                    {
                        exist = 1;
                        row = rowcount;
                        break;
                    }
                  
                }
                if (exist == 0)
                {
                    Sheet.Cells[Rng.Rows.Count+1, 12] = DateTime.Today.ToString();
                    row = Rng.Rows.Count+1;
                }
                WB.Close(false, null , null);                          //ftiaxnw ta pedia sto report

                XL.Quit();
             //   releaseObject(Rng);
                releaseObject(Sheet);
                releaseObject(WB);
                releaseObject(XL);
             //   WB.SaveCopyAs(monthpath + "/dokimastiko.xls");
               // WB.Close(false, null, null);                          //ftiaxnw ta pedia sto report

                //XL.Quit();
            }
            else
            {
                
                
                XL = new Excel.Application();
              //  XL.Visible = false;
                System.IO.File.Copy("c:/Users/dimitris/Dropbox/WhoWorks/records.xls", dokimastiko);
                WB = XL.Workbooks.Open(dokimastiko, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Sheet = (Excel.Worksheet)WB.Worksheets.get_Item(1);
                
               // Convert.ToString((Rng.Cells[1, 30] as Excel.Range).Value2);
                Sheet.Cells[1, 12] = DateTime.Today.ToString();
                Rng = Sheet.UsedRange;
                row = 1;
                WB.SaveCopyAs(monthpath + "/"+monthname+" report.xls");
                //WB.Close(false, null, null);                          //ftiaxnw ta pedia sto report

                //XL.Quit();
                WB.Close(false, null, null);                          //ftiaxnw ta pedia sto report

                XL.Quit();
               // releaseObject(Rng);
                releaseObject(Sheet);
                releaseObject(WB);
                releaseObject(XL);
            }
          //  WB.Close(true, null, null);                          //ftiaxnw ta pedia sto report

            //XL.Quit();
            int x;



            try
            {
             
                string targetPath = monthpath;

                // Use Path class to manipulate file and directory paths. 
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, fileName2);
                //Start Excel and get Application object.
                if (System.IO.File.Exists(destFile))
                {

                    //MessageBox.Show("to arxeio iparxei");
                    sourceFile = destFile;

                }

                else
                {
                    //oWB = oXL.Workbooks.Open(destFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    // MessageBox.Show("laalallala");

                    //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                    //oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    //string name = "omerta";
                    //DateTime today = DateTime.Today;
                    name = today.ToString("D");
                    System.IO.File.Copy(sourceFile, destFile);
                    
                }
                  //  oXL = new Excel.Application();
                   // oXL.Visible = false;
                    //oWB = oXL.Workbooks.Open(destFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    //  oWB.SaveCopyAs("c:/Program Files/WhoWorks/"+name+" report.xls");
                    //oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item(1);
                    
                   // MessageBox.Show("allaallalala");
                



              



              
                 //oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");


                
                //<<<<<<<<<<<<< OPEN record.xls>>>>>>>>>>>>>>>>>>
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(sourceFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                
                range = xlWorkSheet.UsedRange;



               // MessageBox.Show("allaallalala");
                
                for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    employment k = new employment();
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                       
                        str = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                       // MessageBox.Show("prin");
                     
                        //MessageBox.Show("meta");
                        // MessageBox.Show(str);
                        if (cCnt == 1)
                        {
                            k.id = str;

                        }
                        else if (cCnt == 2)
                        {
                            k.namesurname = str;

                        }
                     

                        else if (cCnt == 5)
                        {
                           // MessageBox.Show("allaallalala");
                
                            k.timein = str;
                            if (str == null)
                            {
                                k.inside = 1;
                               // MessageBox.Show("hsshfgdgdksgfdgdls");
                            }
                        }
                        else if (cCnt == 6)
                        {
                            k.timeout = str;
                        }
                        else if (cCnt == 7)
                        {
                            k.numberofbreaks = Convert.ToInt32(str);
                        }
                       
                     
                       
                        else if (cCnt == 12)
                        {
                            if (str == null)
                            {
                                k.inside = 1;
                                // MessageBox.Show("hsshfgdgdksgfdgdls");
                            }
                            else
                            k.inside = Convert.ToInt32(str);
                            
                        }
                        else if (cCnt == 13)
                        {
                            if (str != null)
                            {
                                double kati = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                // k.fullworksec = Convert.ToDouble(str);
                                k.fullworksec = kati;
                                k.fullworktime = TimeSpan.FromSeconds(kati);
                            }
                        
                        }
                        else if (cCnt == 14)
                        {
                            if (str != null)
                            {
                                double kati = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                //k.fullbreaksec = Convert.ToDouble(str);
                                k.fullbreaksec = kati;
                                k.fullbreaktime = TimeSpan.FromSeconds(kati);
                            }
                        }
                        else if (cCnt == 15)
                        {
                            if (str != null)
                            {
                                double kati = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                //k.fullbreaksec = Convert.ToDouble(str);
                                k.starttime = TimeSpan.FromSeconds(kati);
                                k.starttimesec = kati;
                            }
                        }
                        else if (cCnt == 16)
                        {
                            if (str != null)
                            {
                                double kati = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                //k.fullbreaksec = Convert.ToDouble(str);
                                k.breaktime = TimeSpan.FromSeconds(kati);
                                k.breaktimesec = kati;
                            }
                        }
                        else if (cCnt == 17)
                        {
                            if (str != null)
                            {
                                double kati = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                //k.fullbreaksec = Convert.ToDouble(str);
                                k.lasttime = TimeSpan.FromSeconds(kati);
                                k.lasttimesec=kati;
                            }
                        }
                        
                    }
                  // MessageBox.Show("allaallalala");
                
                   // Convert.ToDateTime(k.timein);
                    //Convert.ToDateTime(k.timeout);
                  
                    ListViewItem lvi = new ListViewItem(k.namesurname);
                    lvi.SubItems.Add(k.timein);
                    lvi.SubItems.Add(k.fullworktime.ToString());//Hours + ":" + list[count].fullworktime.Minutes + ":" + list[count].fullworktime.Seconds);
                    lvi.SubItems.Add(k.fullbreaktime.ToString());//.Hours + ":" + list[count].fullbreaktime.Minutes + ":" + list[count].fullbreaktime.Seconds);
                    lvi.SubItems.Add(k.timeout);
                    lvi.SubItems.Add(k.numberofbreaks.ToString());
                    if (k.inside % 2 == 0)
                    {
                        lvi.ForeColor = Color.Green;
                        
                    }
                    else
                        lvi.ForeColor = Color.Red;
                    if (k.inside!=1)
                        k.inside++;
                   // k.inside++;
               
                    listView1.Items.Add(lvi);
                  
                    k.indexoflistbox = listView1.Items.Count - 1;
                    

                    list.Add(k);


                }

               // MessageBox.Show(":akakakakakakakaka");
                xlWorkBook.Close(false, null, null);                          //ftiaxnw ta pedia sto report

                xlApp.Quit();
                //releaseObject(range);
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

            //    MessageBox.Show("ekleise i glika");
                //oWB.Close(true, null, null );                          //ftiaxnw ta pedia sto report

         //       oXL.Quit();
                //releaseObject(oRng);
           //     releaseObject(oSheet);
             //   releaseObject(oWB);
               // releaseObject(oXL);
                
             /*  oSheet.Cells[1, 1] = "ID";
                oSheet.Cells[1, 2] = "ΟΝΟΜΑ/ΕΠΩΝΥΜΟ";
                oSheet.Cells[1, 3] = "ΗΜΕΡΟΜΗΝΙΑ";
                oSheet.Cells[1, 4] = "WORKING HOURS";
                oSheet.Cells[1, 5] = "ΑΦΙΞΗ ";
                oSheet.Cells[1, 6] = "ΑΝΑΧΩΡΗΣΗ";
                oSheet.Cells[1, 7] = "ΑΡΙΘΜΟΣ ΔΙΑΛΕΙΜΜΑΤΩΝ";
                oSheet.Cells[1, 8] = "OVERTIME";
                oSheet.Cells[1, 9] = "BREAK TIME";*/
                 




               // oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");
                
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
          
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

    
        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            //MessageBox.Show("")
            
            //oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");
            string name = "omerta";
           
            DateTime today = DateTime.Today;
            name = today.ToString("D");
            string month = today.ToString("MMM");
            string fileName2 = name+" report.xls";
            string fileName3 =  month+" report.xls";
            string targetPath = "c:/Users/dimitris/Dropbox/WhoWorks/" + month;
            string destFile = System.IO.Path.Combine(targetPath, fileName2);
            string destFiledokimastiko = System.IO.Path.Combine(targetPath, fileName3);
            string value = "08:00:00.00";
            //string dokimastiko = targetPath + "/dokimastiko.xls";
            TimeSpan overtime;
            TimeSpan ko = TimeSpan.Parse(value);
             //XL = new Excel.Application();
               // XL.Visible = false;
               // System.IO.File.Copy("c:/Program Files/WhoWorks/records.xls",dokimastiko);
                //WB = XL.Workbooks.Open(dokimastiko, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //Sheet = (Excel.Worksheet)WB.Worksheets.get_Item(1);
               // Convert.ToString((Rng.Cells[1, 30] as Excel.Range).Value2);
                //Sheet.Cells[1, 12] = DateTime.Today.ToString();
                //Rng = Sheet.UsedRange;
                //row = 1;
                
         
               
            if (e.KeyCode == Keys.Enter)
            {
                //   textBox1.Text = "lalalala";
                Excel.Application oXL;
                Excel._Workbook oWB;                           //orizw tis excel metavlites gia to report
                Excel._Worksheet oSheet;
                Excel.Range oRng;
                Excel.Application XL;
                Excel._Workbook WB;
                Excel._Worksheet Sheet;                        //orizw tis excel metavlites gia to miniaio report;
                Excel.Range Rng;
                XL = new Excel.Application();
                WB = XL.Workbooks.Open(destFiledokimastiko, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Sheet = (Excel.Worksheet)WB.Worksheets.get_Item(1);

                Rng = Sheet.UsedRange;
                oXL = new Excel.Application();
                oXL.Visible = false;
                oWB = oXL.Workbooks.Open(destFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                oWB.SaveCopyAs(destFile);
                oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item(1);
                int count;
                e.Handled = false;


                for (count = 0; count <= list.Count - 1; count++)
                {

                    if (textBox1.Text == list[count].id)
                    {

                        //MessageBox.Show("lkljkhk");
                        int flag = 0;
                        TimeSpan timeSpan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0);
                        if (textBox1.Text == list[count].id && list[count].inside == 1) //enarksi vardias
                        {
                           // MessageBox.Show("mpainei gia prwth fora");
                            list[count].starttime = timeSpan;
                            list[count].starttimesec = timeSpan.TotalSeconds;
                           // MessageBox.Show(timeSpan.ToString());
                            list[count].lasttime = timeSpan;
                            list[count].worktimesec = timeSpan.TotalSeconds;
                          
                            list[count].timein = DateTime.Now.ToLongTimeString();
                            list[count].inside = 2;

                        }
                        else if (textBox1.Text == list[count].id && list[count].inside >= 1 && list[count].inside % 2 == 1)//bgainei gia dialeima
                        {
                           // MessageBox.Show("vgainei gia dialeima");
                            list[count].timeout = DateTime.Now.ToLongTimeString();
                            list[count].breaktime = timeSpan;
                            list[count].breaktimesec = timeSpan.TotalSeconds;
                           // list[count].lasttime = timeSpan;
                           
                            overtime = list[count].breaktime - list[count].starttime;
                             list[count].fullworktime  += (timeSpan - list[count].lasttime);
                        list[count].fullworksec = list[count].fullworktime.TotalSeconds;
                            if (overtime.TotalSeconds >=ko.TotalSeconds)
                                list[count].overtime = overtime - ko;


                        }
                        else if (textBox1.Text == list[count].id && list[count].inside >= 1 && list[count].inside % 2 == 0)//mpainei gia douleia
                        {
                           // MessageBox.Show("mpainei gia douleia");
                            list[count].fullbreaktime += (timeSpan - list[count].breaktime);
                            list[count].lasttime = timeSpan;
                            list[count].worktimesec = timeSpan.TotalSeconds;
                            list[count].fullbreaksec = list[count].fullbreaktime.TotalSeconds;
                            list[count].numberofbreaks++;

                        }

                       // list[count].fullworktime = timeSpan - list[count].starttime - list[count].fullbreaktime;
                      //  fullworksec = list[count].fullworktime.TotalSeconds;
                        //int t = list[count].fullworktime.Hours;
                        //int k = list[count].fullworktime.Minutes;
                        //int a = list[count].fullworktime.Seconds;
                       
                        
                        //TimeSpan kra = TimeSpan.FromSeconds(tsou);
                      
                       // list[count].fullworktime.Hours = t;
                       // MessageBox.Show(list[count].fullworktime.ToString());
                       // MessageBox.Show(t.ToString());
                        //MessageBox.Show(k.ToString());
                        //MessageBox.Show(a.ToString());
                       // MessageBox.Show(kra.ToString());
                       
                        ListViewItem lvi = new ListViewItem(list[count].namesurname);

                        
                        oSheet.Cells[(list[count].indexoflistbox + 2), 1] = list[count].id;
                      
                        oSheet.Cells[(list[count].indexoflistbox + 2), 2] = list[count].namesurname;


                        oSheet.Cells[(list[count].indexoflistbox + 2), 4] = list[count].fullworktime.Hours + ":" + list[count].fullworktime.Minutes + ":" + list[count].fullworktime.Seconds;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 5] = list[count].timein;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 6] = list[count].timeout;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 7] = list[count].numberofbreaks;
                        TimeSpan time = TimeSpan.Zero;
                        if (list[count].overtime != time)
                            oSheet.Cells[(list[count].indexoflistbox + 2), 8] = list[count].overtime.Hours + ":" + list[count].overtime.Minutes + ":" + list[count].overtime.Seconds;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 9] = list[count].fullbreaktime.Hours + ":" + list[count].fullbreaktime.Minutes + ":" + list[count].fullbreaktime.Seconds;

                        oSheet.Cells[(list[count].indexoflistbox + 2), 3] = DateTime.Today.ToString();
                      //  oSheet.Cells[(list[count].indexoflistbox + 2), 10] = list[count].starttime.ToString();
                       // oSheet.Cells[(list[count].indexoflistbox + 2), 11] = list[count].breaktime.ToString();
                        oSheet.Cells[(list[count].indexoflistbox + 2), 12] = list[count].inside.ToString();
                        oSheet.Cells[(list[count].indexoflistbox + 2), 13] = list[count].fullworksec;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 14] = list[count].fullbreaksec;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 15] = list[count].starttimesec;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 16] = list[count].breaktimesec; 
                           oSheet.Cells[(list[count].indexoflistbox + 2), 17] = list[count].worktimesec; 


                        

                        oWB.SaveCopyAs(targetPath+"/"+name+" report.xls");



                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 1] = list[count].id;

                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 2] = list[count].namesurname;


                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 4] = list[count].fullworktime.Hours + ":" + list[count].fullworktime.Minutes + ":" + list[count].fullworktime.Seconds;
                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 5] = list[count].timein;
                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 6] = list[count].timeout;
                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 7] = list[count].numberofbreaks;
                        Sheet.Cells[(list[count].indexoflistbox + 2 + row), 8] = list[count].overtime.ToString();
                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 9] = list[count].fullbreaktime.Hours + ":" + list[count].fullbreaktime.Minutes + ":" + list[count].fullbreaktime.Seconds;

                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 3] = DateTime.Today.ToString();
                        //Sheet.Cells[(list[count].indexoflistbox + 2+row), 10] = list[count].starttime.ToString();
                       // Sheet.Cells[(list[count].indexoflistbox + 2+row), 11] = list[count].breaktime.ToString();
                        Sheet.Cells[(list[count].indexoflistbox + 2+row), 12] = list[count].inside.ToString();




                        WB.SaveCopyAs(targetPath+"/"+month+" report.xls");
                        countexcel++;

                        lvi.SubItems.Add(list[count].timein);
                        lvi.SubItems.Add(list[count].fullworktime.ToString());//Hours + ":" + list[count].fullworktime.Minutes + ":" + list[count].fullworktime.Seconds);
                        lvi.SubItems.Add(list[count].fullbreaktime.ToString());//.Hours + ":" + list[count].fullbreaktime.Minutes + ":" + list[count].fullbreaktime.Seconds);
                        lvi.SubItems.Add(list[count].timeout);
                        lvi.SubItems.Add(list[count].numberofbreaks.ToString());

                        if (list[count].inside % 2 == 0 || list[count].inside==1)
                        {
                            lvi.ForeColor = Color.Green;
                        }
                        else
                            lvi.ForeColor = Color.Red;

                        listView1.Items[list[count].indexoflistbox] = lvi;

                        list[count].inside++;
                        break;
                    }
                    else if (count == list.Count - 1)
                    {
                        MessageBox.Show("dn iparxei o ipallilos");
                    }
                }

                textBox1.Text = String.Empty;
                employment x;
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = true;
                oWB.Close(false, null, null );                          //ftiaxnw ta pedia sto report

                       oXL.Quit();
               // releaseObject(oRng);
                     releaseObject(oSheet);
                   releaseObject(oWB);
                 releaseObject(oXL);
                 WB.Close(false, null, null);                          //ftiaxnw ta pedia sto report

                 XL.Quit();
                 // releaseObject(Rng);
                 releaseObject(Sheet);
                 releaseObject(WB);
                 releaseObject(XL);
                
            }
           // Marshal.
            //WB.Close(false, null, null);                          //ftiaxnw ta pedia sto report
            
            //XL.Quit();
            //Marshal.ReleaseComObject(sheets);
            //Marshal.ReleaseComObject(sheet);
         //   foreach (Process process in Process.GetProcessesByName("Excel"))
           // {
             //   process.Kill();
            //}

           
           
            
        }
        
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            
         //   gHook = new GlobalKeyboardHook();                                   //tin kalw opws akrivwss sto youtube
           // gHook.KeyDown += new KeyEventHandler(gHook_KeyDown);
            // Add the keys you want to hook to the HookedKeys list
            //foreach (Keys key in Enum.GetValues(typeof(Keys)))
             //   gHook.HookedKeys.Add(key);
            
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }


       
                               //ftiaxnw ta pedia sto report
                
              

    }
     
               
}
