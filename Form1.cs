using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Schedule
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Get_Schedule_Click(object sender, EventArgs e)
        {

            OutlookManager outlookManager = new OutlookManager();

            // メールアドレス もしくは 名前 を指定
            // 対象日時を指定（Start～Endでの指定もできます）

            //long Kikan =  long.Parse(textBox_kikan);

            string Kikan_text = textBox_kikan.Text;
            int Kikan = Convert.ToInt16(Kikan_text);

            var schedules = outlookManager.GetScheduleList(textBox_User1+"@pc-daiwabo.co.jp", DateTime.Today,Kikan);

            //ファイルの初期化 
            File.Delete(@"c:\temp\schedule.txt");

            foreach (var s in schedules)
                //    Console.WriteLine(s);


                //File.AppendAllText(@"c:\temp\schedule.txt", Convert.ToString(s) + " " +Convert.ToString(s.Start.DayOfWeek) + Environment.NewLine);
                //File.AppendAllText(@"c:\temp\schedule.txt", Convert.ToString(s.Start.Date) + Convert.ToString(s.Start.Hour) + " ～ " + Convert.ToString(s.End.Hour) +"("+ Convert.ToString(s.Start.DayOfWeek) + ") " + Convert.ToString(s.Title) + Environment.NewLine);
                File.AppendAllText(@"c:\temp\schedule.txt",  Convert.ToString(s) + "  (" + Convert.ToString(s.Start.DayOfWeek.ToString()) + ") " + Environment.NewLine);
            
                
               // File.AppendAllText(@"c:\temp\schedule.txt", "(" + Convert.ToString(s.Start.DayOfWeek.ToString("d")) +  ") " + Environment.NewLine + "  "  + Convert.ToString(s)  + Environment.NewLine);

               // File.AppendAllText(@"c:\temp\schedule.txt",Convert.ToString(s.Start.DayOfWeek.ToString("d")) + ") " + Environment.NewLine + "  " + Convert.ToString(s) + Environment.NewLine);  

            //File.AppendAllText(@"c:\temp\schedule.txt", "(" + Convert.ToString(("日月火水木金土").Substring(int.Parse(s.Start.ToString("d")), 1)) + ")"  + Convert.ToString(s) + Environment.NewLine);



            // File.AppendAllText(@"c:\temp\schedule.txt", "(" + Convert.ToString(s.Start.DayOfWeek.ToString()) + ") " + Convert.ToString(s) + Environment.NewLine);


            //テキストを開く
            System.Diagnostics.Process.Start("notepad.exe", @"""C:\temp\schedule.txt""");

            //    t = t +" "+ Convert.ToString(s);                
            //https://qiita.com/t_sato310/items/6c6ec79f4487a52b6b9e
            //File.WriteAllText(@"c:\temp\schedule.txt",t);
            // MessageBox.Show(t);   

            // Console.WriteLine($"Schedules count={schedules.Count}");

        }


    }
}
