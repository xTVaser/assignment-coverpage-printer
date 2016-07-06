using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Fonts;
using System.Diagnostics;
using MigraDoc.DocumentObjectModel.Tables;

namespace Assignment_Coverpage_Printer {
    class Program {

        [STAThread]
        static void Main(string[] args) {

            List<string> students = collectNames();

            List<string> assignmentInfo = collectMetaInfo();

            List<string> assignmentMarking = collectRubrik();

            exportPDF(students, assignmentInfo, assignmentMarking);
        }

        /// <summary>
        /// <exception cref="FileNotFoundException">File was not found!</exception>
        /// </summary>
        /// <returns></returns>
        static List<string> collectNames() {

            OpenFileDialog fb = new OpenFileDialog();
            fb.Filter = "Text Files (*.txt)|*.txt";
            DialogResult folder = fb.ShowDialog();


            //Safe FileName = name, FileName = Path
            Console.WriteLine(fb.FileName);

            return new List<string>(File.ReadAllLines(fb.FileName));
        }

        static List<string> collectMetaInfo() {

            SimpleForm metadata = new SimpleForm();
            return new List<string>(metadata.ShowDialog());
        }

        static List<string> collectRubrik() {

            Console.WriteLine("How Many Questions are in this Assignment?: ");
            int numAssignments = Convert.ToInt32(Console.ReadLine());

            List<string> rubrik = new List<string>();

            Console.WriteLine("What percentage are the programming questions worth? (Out of 100): ");
            rubrik.Add(Console.ReadLine());
            Console.WriteLine("What percentage is the demonstration worth? (Out of 100): ");
            rubrik.Add(Console.ReadLine());
            Console.WriteLine("What percentage is the comments deduction? (Out of 100): ");
            rubrik.Add(Console.ReadLine());
            Console.WriteLine("What percentage is the formatting deduction? (Out of 100): ");
            rubrik.Add(Console.ReadLine());
            Console.WriteLine("What percentage is the late penalty? (Out of 100): ");
            rubrik.Add(Console.ReadLine());

            for(int i = 0; i < numAssignments; i++) {

                Console.WriteLine("Question " + (i + 1) + " is called?: ");
                rubrik.Add(Console.ReadLine());
                Console.WriteLine("Question " + (i + 1) + " is out of?: ");
                rubrik.Add(Console.ReadLine());
            }

            return rubrik;
        }

        static void exportPDF(List<string> students, List<string> info, List<string> rubrik) {

            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "PDF File (*.pdf|*.pdf";

            if (saveFile.ShowDialog() != DialogResult.OK)
                return;

            Document doc = new Document();

            string assignmentName = info[0];
            string courseCode = info[1];
            string prof = info[2];
            string ta = info[3];

            //MetaData
            doc.Info.Title = assignmentName + " Cover Page Print-Outs " + courseCode;
            doc.Info.Author = "Tyler Wilding";
            
            //Document Styles
            Style style = doc.Styles["Normal"];
            style.Font.Name = "Segoe UI";
            style.Font.Size = 12;

            style = doc.Styles["Heading1"];
            style.Font.Name = "Segoe UI";
            style.Font.Size = 24;
            style.Font.Color = Colors.Black;
            style.Font.Bold = true;

            style = doc.Styles["Heading2"];
            style.Font.Name = "Segoe UI";
            style.Font.Size = 18;
            style.Font.Color = Colors.DarkMagenta;

            style = doc.Styles.AddStyle("Small", "Normal");
            style.Font.Name = "Segoe UI";
            style.Font.Size = 10;
            style.Font.Color = Colors.Black;
            style.Font.Italic = true;

            foreach (string student in students) {

                //Front Cover
                Section section = doc.AddSection();
                Paragraph paragraph = section.AddParagraph(student);
                paragraph.Format.SpaceAfter = "3.5cm";
                paragraph.Format.Alignment = ParagraphAlignment.Center;
                paragraph.Style = "Heading1";

                paragraph = section.AddParagraph(courseCode);
                paragraph.Format.SpaceAfter = "2.5cm";
                paragraph.Format.Alignment = ParagraphAlignment.Center;
                paragraph.Style = "Heading2";

                paragraph = section.AddParagraph(assignmentName);
                paragraph.Format.SpaceAfter = "2.5cm";
                paragraph.Format.Alignment = ParagraphAlignment.Center;
                paragraph.Style = "Heading2";

                paragraph = section.AddParagraph("Professor: " + prof);
                paragraph.Format.SpaceAfter = "2.5cm";
                paragraph.Format.Alignment = ParagraphAlignment.Center;
                paragraph.Style = "Heading2";

                paragraph = section.AddParagraph("Teaching Assistant: " + ta);
                paragraph.Format.SpaceAfter = "2.5cm";
                paragraph.Format.Alignment = ParagraphAlignment.Center;
                paragraph.Style = "Heading2";

                section.AddPageBreak();

                //Marking Breakdown
                paragraph = section.AddParagraph("Marking Breakdown");
                paragraph.Format.Alignment = ParagraphAlignment.Center;
                paragraph.Format.SpaceAfter = "0.25cm";
                paragraph.Style = "Heading2";

                Table table = new Table();
                table.Rows.Alignment = RowAlignment.Center;
                table.Borders.Width = 0.75;

                table.AddColumn(Unit.FromCentimeter(8));

                Column column = table.AddColumn(Unit.FromCentimeter(2));
                column.Format.Alignment = ParagraphAlignment.Right;

                Row row = table.AddRow();
                row.Shading.Color = Colors.PaleGreen;
                row.Style = "Normal";
                Cell cell = row.Cells[0];
                paragraph = section.AddParagraph();
                cell.AddParagraph("Programming Questions");
                cell = row.Cells[1];
                cell.AddParagraph(rubrik[0] + "%");
                
                row = table.AddRow();
                row.Shading.Color = Colors.PaleGreen;
                row.Style = "Normal";
                cell = row.Cells[0];
                cell.AddParagraph("Demonstration");
                cell = row.Cells[1];
                cell.AddParagraph(rubrik[1] + "%");
                
                row = table.AddRow();
                row.Shading.Color = Colors.PaleVioletRed;
                row.Style = "Normal";
                cell = row.Cells[0];
                cell.AddParagraph("Comment Deduction");
                cell = row.Cells[1];
                cell.AddParagraph(rubrik[2] + "%");

                row = table.AddRow();
                row.Shading.Color = Colors.PaleVioletRed;
                row.Style = "Normal";
                cell = row.Cells[0];
                cell.AddParagraph("Formatting Deduction");
                cell = row.Cells[1];
                cell.AddParagraph(rubrik[3] + "%");

                row = table.AddRow();
                row.Shading.Color = Colors.PaleVioletRed;
                row.Style = "Normal";
                cell = row.Cells[0];
                cell.AddParagraph("Late Penalty");
                cell = row.Cells[1];
                cell.AddParagraph(rubrik[4] + "%");

                table.SetEdge(0, 0, 2, 3, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 1.5, Colors.Black);
                table.Format.SpaceAfter = "1.5cm";

                section.Add(table);

                //Programming Problems
                paragraph = section.AddParagraph("Programming Problems");
                paragraph.Style = "Heading1";
                paragraph.Format.SpaceAfter = "0.25cm";

                for(int i = 5; i < rubrik.Count; i+=2) {

                    paragraph = section.AddParagraph(rubrik[i]+" (\t\t/"+rubrik[i+1]+")");
                    paragraph.Style = "Heading2";
                    paragraph.Format.LeftIndent = "1cm";
                    
                    paragraph = section.AddParagraph("Notes: ");
                    paragraph.Style = "Normal";
                    paragraph.Format.SpaceAfter = "1cm";
                    paragraph.Format.LeftIndent = "2cm";
                }
                
                paragraph = section.AddParagraph("Demonstration " + "(\t\t/" + rubrik[1]+")");
                paragraph.Format.Alignment = ParagraphAlignment.Left;
                paragraph.Style = "Heading1";

                paragraph = section.AddParagraph("Notes: ");
                paragraph.Style = "Normal";
                paragraph.Format.SpaceAfter = "2cm";
                paragraph.Format.LeftIndent = "1cm";
            
                paragraph = section.AddParagraph("Comment Deduction " + "(\t\t/" + rubrik[2] + ")");
                paragraph.Format.Alignment = ParagraphAlignment.Left;
                paragraph.Style = "Heading1";

                paragraph = section.AddParagraph("Notes: ");
                paragraph.Style = "Normal";
                paragraph.Format.SpaceAfter = "2cm";
                paragraph.Format.LeftIndent = "1cm";

                paragraph = section.AddParagraph("Formatting Deduction " + "(\t\t/" + rubrik[3] + ")");
                paragraph.Format.Alignment = ParagraphAlignment.Left;
                paragraph.Style = "Heading1";

                paragraph = section.AddParagraph("Notes: ");
                paragraph.Style = "Normal";
                paragraph.Format.SpaceAfter = "2cm";
                paragraph.Format.LeftIndent = "1cm";
            }

            //Setting Up Exporting
            MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(doc, "MigraDoc.mddd1");
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.Always);
            renderer.Document = doc;
            renderer.RenderDocument();

            //Name of document
            string filename = saveFile.FileName;
            renderer.PdfDocument.Save(filename);
            Process.Start(filename);
        }
    }
}
