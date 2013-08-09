
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

        GlobalKeyboardHook gHook; 
         
        Excel.Application xlApp;                 //arxikopoiw tin excel vivliothiki
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        List<employment> list = new List<employment>();
        String str;
        int rCnt = 0;
        int cCnt = 0;
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;
        Excel.Range oRng;
        int countexcel=2;
         






        class employment
        {
           
            public string id;
            public string namesurname;
          
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
                namesurname = null;
                
                inside = 0;
                numberofbreaks = 0;
                //starttime = DateTime.
              //  breaktime= 0;
                //fullbreaktime = 0;

            }
        }
        public Form1()
        {
            //this.WindowState = FormWindowState.Minimized;
           // this.ShowInTaskbar = true;
            InitializeComponent();
          
            string fileName = "records.xls";
            string fileName2 = "report.xls";
            string sourcePath = "c:/Program Files/WhoWorks";
            

            int x;



            try
            {
               
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                 oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");







                string targetPath = "c:/Program Files/WhoWorks";

                // Use Path class to manipulate file and directory paths. 
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, fileName2);



              
                 oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");


                

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
                            k.namesurname = str;

                        }


                    }

                  
                    ListViewItem lvi = new ListViewItem(k.namesurname);
                    lvi.ForeColor = Color.Red;
               
                    listView1.Items.Add(lvi);
                  
                    k.indexoflistbox = listView1.Items.Count - 1;
                    

                    list.Add(k);


                }



               
                xlWorkBook.Close(true, null, null);                          //ftiaxnw ta pedia sto report
                xlApp.Quit();
                oSheet.Cells[1, 1] = "ID";
                oSheet.Cells[1, 2] = "ΟΝΟΜΑ/ΕΠΩΝΥΜΟ";
                oSheet.Cells[1, 3] = "ΗΜΕΡΟΜΗΝΙΑ";
                oSheet.Cells[1, 4] = "WORKING HOURS";
                oSheet.Cells[1, 5] = "ΑΦΙΞΗ ";
                oSheet.Cells[1, 6] = "ΑΝΑΧΩΡΗΣΗ";
                oSheet.Cells[1, 7] = "ΑΡΙΘΜΟΣ ΔΙΑΛΕΙΜΜΑΤΩΝ";
                oSheet.Cells[1, 8] = "OVERTIME";
                oSheet.Cells[1, 9] = "BREAK TIME";




                oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");
                
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

        public void gHook_KeyDown(object sender, KeyEventArgs e)        //i sinartisi gia ta hook kaleite sto form load apo katw 
        {
           System.Diagnostics.Debugger.Break();
           
            if (e.KeyCode != Keys.Enter)
            textBox1.Text += ((char)e.KeyValue).ToString();
           if (e.KeyCode == Keys.Enter)
            {

                int count;
                e.Handled = true;

              
                for (count = 0; count <= list.Count - 1; count++)
                {
                    
                    if (textBox1.Text == list[count].id)
                    {
                        MessageBox.Show("lkljkhk");
                        int flag = 0;
                        TimeSpan timeSpan = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0);
                        if (textBox1.Text == list[count].id && list[count].inside == 0) //enarksi vardias
                        {
                           
                            list[count].starttime = timeSpan;
                            list[count].timein = DateTime.Now.ToLongTimeString();
                          
                        }
                        else if (textBox1.Text == list[count].id && list[count].inside >= 1 && list[count].inside % 2 == 1)//bgainei gia dialeima
                        {
                            
                            list[count].timeout = DateTime.Now.ToLongTimeString();
                            list[count].breaktime = timeSpan;
                            list[count].lasttime = timeSpan;
                            list[count].numberofbreaks++;
                         


                        }
                        else if (textBox1.Text == list[count].id && list[count].inside >= 1 && list[count].inside % 2 == 0)//mpainei gia douleia
                        {
                            
                            list[count].fullbreaktime += (timeSpan - list[count].breaktime);
                           



                        }

                        list[count].fullworktime = timeSpan - list[count].starttime - list[count].fullbreaktime;
                        ListViewItem lvi = new ListViewItem(list[count].namesurname);
                    

                        oSheet.Cells[(list[count].indexoflistbox + 2), 1] = list[count].id;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 2] = list[count].namesurname;

                       
                        oSheet.Cells[(list[count].indexoflistbox + 2), 4] = list[count].fullworktime.Hours + ":" + list[count].fullworktime.Minutes + ":" + list[count].fullworktime.Seconds;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 5] = list[count].timein;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 6] = list[count].timeout;
                        oSheet.Cells[(list[count].indexoflistbox + 2), 7] = list[count].numberofbreaks;

                        oSheet.Cells[(list[count].indexoflistbox + 2), 9] = list[count].fullbreaktime.Hours + ":" + list[count].fullbreaktime.Minutes + ":" + list[count].fullbreaktime.Seconds;

                        oSheet.Cells[(list[count].indexoflistbox + 2), 3] = DateTime.Today.ToString();
                        

                        oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");
                        countexcel++;

                        lvi.SubItems.Add(list[count].timein);
                        lvi.SubItems.Add(list[count].fullworktime.Hours + ":" + list[count].fullworktime.Minutes + ":" + list[count].fullworktime.Seconds);
                        lvi.SubItems.Add(list[count].fullbreaktime.Hours + ":" + list[count].fullbreaktime.Minutes + ":" + list[count].fullbreaktime.Seconds);
                        lvi.SubItems.Add(list[count].timeout);
                        lvi.SubItems.Add(list[count].numberofbreaks.ToString());
                      
                        if (list[count].inside % 2 == 0)
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
             
            }
           
        }
        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {


           
            oWB.SaveCopyAs("c:/Program Files/WhoWorks/report.xls");

            string fileName2 = "report.xls";
            string targetPath = "c:/Program Files/WhoWorks";
            string destFile = System.IO.Path.Combine(targetPath, fileName2);
           
           
           
            
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
            gHook = new GlobalKeyboardHook();                                   //tin kalw opws akrivwss sto youtube
            gHook.KeyDown += new KeyEventHandler(gHook_KeyDown);
            // Add the keys you want to hook to the HookedKeys list
            foreach (Keys key in Enum.GetValues(typeof(Keys)))
                gHook.HookedKeys.Add(key);
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            gHook.hook();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            gHook.unhook();
        }


       


    }
  
}

