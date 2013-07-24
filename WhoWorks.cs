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

namespace WindowsFormsApplication15
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        List<employment> list = new List<employment>();
        String str;
        int rCnt = 0;
        int cCnt = 0;






        class employment
        {

            public string id;
            public string name;
            public string surname;
            public string timein;
            public string timeout;
            public int inside;  // 1==in     0==out
            public int indexoflistbox;
            public int numberofbreaks;
            //public double breaktime;
            //public double fullbreaktime;
            public TimeSpan starttime;
            public TimeSpan breaktime;
            public TimeSpan fullbreaktime;
            public TimeSpan fullworktime;
            public TimeSpan lasttime;
            public employment()
            {
                id = null;
                name = null;
                surname = null;
                inside = 0;
                numberofbreaks = 0;
                //starttime = DateTime.
              //  breaktime= 0;
                //fullbreaktime = 0;

            }
        }
        public Form1()
        {
            InitializeComponent();
            
            string fileName = "records.xls";
            string fileName2 = "report.xls";
            string sourcePath = "c:/Program Files/WhoWorks";
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            int x;



            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                //oWB.SaveCopyAs("C:/Users/dimitris/Desktop/first project/report.xls");





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
            string targetPath = "c:/Program Files/WhoWorks";

            // Use Path class to manipulate file and directory paths. 
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName2);


            //System.IO.File.Copy(sourceFile, destFile, true);

            /*
                         oXL = new Excel.Application();
                           oWB = oXL.Workbooks.Open(sourceFile);
                           oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item(1);
                           oWB.Close(true, null, null);
                           oRng = oSheet.UsedRange;
                           for (rCnt = 1; rCnt <= oRng.Rows.Count; rCnt++)
                           {
                               for (cCnt = 1; cCnt <= oRng.Columns.Count; cCnt++)
                               {
                                   str = (string)(oRng.Cells[rCnt, cCnt] as Excel.Range).Value2;
                                   MessageBox.Show(str);
                               }
                           }
                           oWB.Close(true, null, null);
                           oXL.Quit();
                           releaseObject(oSheet);
                           releaseObject(oWB);
                           releaseObject(oXL);

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
                               MessageBox.Show("Unable to release the Object " + ex.ToString());
                           }
                           finally
                           {
                               GC.Collect();
                           }
                       } 
               
                      */

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sourceFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;




            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                employment k = new employment();
                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {

                    str = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    // MessageBox.Show(str);
                    if (cCnt == 1)
                    {
                        k.id = str;

                    }
                    else if (cCnt == 2)
                    {
                        k.name = str;

                    }
                    else if (cCnt == 3)
                    {
                        k.surname = str;
                    }

                }

                 listBox1.Items.Add (k.name + "    " + k.surname );
                 ListViewItem lvi= new ListViewItem(k.name);
                 lvi.ForeColor = Color.Red;
                 lvi.SubItems.Add(k.surname);
                 listView1.Items.Add(lvi);
                // listView1.Items.Ad
                 //listView1.Items.Add(k.surname); 
                 k.indexoflistbox = listBox1.Items.Count-1;
                 
                list.Add(k);


            }

            //MessageBox.Show(list.Count.ToString());
            //MessageBox.Show(list[1].name);
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            string id;
           // DateTime la;
            if (e.KeyCode == Keys.Enter)
            {
                int count;
                e.Handled = true;

                id = textBox1.Text;
                for (count = 0; count <= list.Count - 1; count++)
                {
                    if (id == list[count].id)
                  {
                      int flag = 0;
                    TimeSpan timeSpan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0);
                    if (id == list[count].id && list[count].inside==0) //enarksi vardias
                    {
                        //MessageBox.Show("mpika");
                        //TimeSpan timeSpan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0);
                        list[count].starttime = timeSpan;
                        list[count].timein = DateTime.Now.ToLongTimeString();
                       //  listBox1.Items[list[count].indexoflistbox]= listBox1.Items[list[count].indexoflistbox]+"\r";
                        //list[count].indexoflistbox = listBox1.Items.Count-1;
                        //list[count].inside++;
                       // la = DateTime.UtcNow;
                        //MessageBox.Show(la.ToString());
                        //TimeSpan t = (DateTime.UtcNow - new DateTime(1970, 1, 1));
                      //  MessageBox.Show(t.TotalSeconds.ToString());
 
                       // break;
                    }
                    else if (id == list[count].id && list[count].inside >= 1 && list[count].inside%2==1)//bgainei gia dialeima
                    {
                        //TimeSpan timeSpan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0);
                        list[count].timeout = DateTime.Now.ToLongTimeString();
                        list[count].breaktime = timeSpan;
                        list[count].lasttime = timeSpan;
                        list[count].numberofbreaks++;
                        //list[count].inside++;
                       //list[count].timeout = DateTime.Now.ToLongTimeString();
                         // listBox1.Items[list[count].indexoflistbox]=  list[count].name + "     " + list[count].surname +"      "+ list[count].timein+"     "+list[count].timeout+"\r";
                          //list[count].inside = 1;
                         // break;
                       

                    }
                    else if (id == list[count].id && list[count].inside >= 1 && list[count].inside % 2 == 0)//mpainei gia douleia
                    {
                       // TimeSpan timeSpan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0);
                        list[count].fullbreaktime += (timeSpan-list[count].breaktime);
                        //list[count].inside++;
                        //MessageBox.Show(list[count].fullbreaktime.Hours.ToString()+" "+list[count].fullbreaktime.Minutes.ToString()+" "+list[count].fullbreaktime.Seconds.ToString()+" ");
                        //break;
                        


                    }

                    list[count].fullworktime = timeSpan - list[count].starttime - list[count].fullbreaktime;
                    ListViewItem lvi = new ListViewItem(list[count].name);
                    lvi.SubItems.Add(list[count].surname);
                    lvi.SubItems.Add(list[count].timein);
                    lvi.SubItems.Add(list[count].fullworktime.Hours + ":" + list[count].fullworktime.Minutes + ":" + list[count].fullworktime.Seconds);
                    lvi.SubItems.Add(list[count].fullbreaktime.Hours + ":" + list[count].fullbreaktime.Minutes + ":" + list[count].fullbreaktime.Seconds);
                    lvi.SubItems.Add(list[count].timeout);
                    lvi.SubItems.Add(list[count].numberofbreaks.ToString());
                   // lvi.SubItems.Add(list[count].numberofbreaks);
                    if (list[count].inside % 2 == 0)
                    {
                        lvi.ForeColor = Color.Green;
                    }
                    else
                        lvi.ForeColor = Color.Red;
                    
                    listView1.Items[list[count].indexoflistbox]=lvi;
                  //  listView1.Items[list[count].indexoflistbox].SubItems[1].
                        listBox1.Items[list[count].indexoflistbox] = list[count].name + "   " + list[count].surname + "        " + list[count].timein + "                      " + list[count].fullworktime.Hours +":"+list[count].fullworktime.Minutes+":"+list[count].fullworktime.Seconds+"             " + list[count].fullbreaktime.Hours+":" +list[count].fullbreaktime.Minutes+":"+list[count].fullbreaktime.Seconds+ "                  " + list[count].timeout + "\r";
                    list[count].inside++;
                    break;
                    }
                else   if (count == list.Count - 1)
                    {
                        MessageBox.Show("dn iparxei o ipallilos");
                    }
                }
                
                textBox1.Text = String.Empty;
                employment x;

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
    }
}
