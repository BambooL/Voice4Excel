using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Speech.Recognition;
using System.Windows;
using System.Diagnostics;
using System.Windows.Forms;
using System.Speech.Synthesis;

//class Sample
//{
//    public static void Main()
//    {
//        Console.WriteLine();
//        //  <-- Keep this information secure! -->
//        Console.WriteLine("SystemDirectory: {0}", Environment.SystemDirectory);
//    }
//}

namespace ExcelAddIn3
{
    public class Word
    {
        public Word() { }
        public string Text { get; set; }
        public string AttachedText { get; set; }
        public bool IsShellCommand { get; set; }
    }
    public partial class ThisAddIn
    {
        public SpeechSynthesizer synth = new SpeechSynthesizer();
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();
            recognizer.SetInputToDefaultAudioDevice();
            synth.SetOutputToDefaultAudioDevice();
            // Create a simple grammar that recognizes "red", "green", or "blue".
            Choices grammar = new Choices();
            grammar.Add(new string[] { "mark", "draw a chart", "get the average", "show me the data", "good" });

            // Create a GrammarBuilder object and append the Choices object.
            GrammarBuilder gb = new GrammarBuilder();
            gb.Append(grammar);

            // Create the Grammar instance and load it into the speech recognition engine.
            Grammar g = new Grammar(gb);
            recognizer.LoadGrammar(g);
            //recognizer.SetInputToDefaultAudioDevice();
            // Register a handler for the SpeechRecognized event.
            recognizer.SpeechRecognized +=
                new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
            recognizer.RecognizeAsync(RecognizeMode.Multiple);


        // Create a simple handler for the SpeechRecognized event.
        
            
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            Console.WriteLine(e.Result.Text);
            if (e.Result.Text == "show me the data") {
                var excelApp = this.Application;
                var workbook = excelApp.ActiveWorkbook;
                excelApp.Visible = true;
                Excel.Worksheet worksheet = ((Excel.Worksheet)Application.ActiveWorkbook.Worksheets[1]);

                worksheet.Cells[1, "B"] = 100;
                worksheet.Cells[2, "B"] = 200;
                worksheet.Cells[3, "B"] = 300;
                worksheet.Cells[4, "B"] = 400;
                worksheet.Cells[5, "B"] = 500;
                worksheet.Cells[6, "B"] = 600;
                worksheet.Cells[7, "B"] = 700;
                worksheet.Cells[8, "B"] = 800;
                worksheet.Cells[1, "A"] = "Sam";
                worksheet.Cells[2, "A"] = "Jack";
                worksheet.Cells[3, "A"] = "Lucy";
                worksheet.Cells[4, "A"] = "Tim";
                worksheet.Cells[5, "A"] = "Alice";
                worksheet.Cells[6, "A"] = "Bob";
                worksheet.Cells[7, "A"] = "Jim";
                worksheet.Cells[8, "A"] = "Andrew";
                synth.Speak("Here I display the data!");
            }

            if (e.Result.Text == "get the average")
            {
                var excelApp = this.Application;
                var workbook = excelApp.ActiveWorkbook;
                excelApp.Visible = true;
                Excel.Worksheet worksheet = ((Excel.Worksheet)Application.ActiveWorkbook.Worksheets[1]);
                worksheet.Cells[9, "A"] = "Average";
                worksheet.Cells[9, "B"].Formula = "=AVERAGE(B1:B8)";
                synth.Speak("Okay, Let's calculate the mean value!");
            }

            if (e.Result.Text == "draw a chart")
            {
                var excelApp = this.Application;
                var workbook = excelApp.ActiveWorkbook;
                excelApp.Visible = true;
                Excel.Worksheet worksheet = ((Excel.Worksheet)Application.ActiveWorkbook.Worksheets[1]);
                var charts = worksheet.ChartObjects() as Excel.ChartObjects;
                var chartObject = charts.Add(60, 10, 300, 300) as Excel.ChartObject;
                var chart = chartObject.Chart;

                // Set chart range.
                var range = worksheet.get_Range("A1", "B8");
                chart.SetSourceData(range);

                // Set chart properties.
                chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xl3DColumn;
                chart.ChartWizard(Source: range,
                    Title: "Gradebook for People",
                    CategoryTitle: "People",
                    ValueTitle: "Grade");
                synth.Speak("Okay, Let's draw a graph!");
            }


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        void DisplayInExcel(IEnumerable<Account> accounts,
           Action<Account, Excel.Range> DisplayFunc)
        {
            var excelApp = this.Application;
            // Add a new Excel workbook.
            excelApp.Workbooks.Add();
            excelApp.Visible = true;
            excelApp.Range["A1"].Value = "ID";
            excelApp.Range["B1"].Value = "Balance";
            excelApp.Range["A2"].Select();

            foreach (var ac in accounts)
            {
                DisplayFunc(ac, excelApp.ActiveCell);
                excelApp.ActiveCell.Offset[1, 0].Select();
            }
            // Copy the results to the Clipboard.
            excelApp.Range["A1:B3"].Copy();
            excelApp.Columns[1].AutoFit();
            excelApp.Columns[2].AutoFit();
        }

 
        #endregion
    }
}
