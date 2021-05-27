using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Word = Microsoft.Office.Interop.Word;


namespace ExamRecog
{
    class Solver
    {
        Word._Application oWord = new Word.Application();
        Word._Document oDoc;
        object oMissing = System.Reflection.Missing.Value;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        List<List<Point>> points = new List<List<Point>>();
        public Solver()
        {
            //Start Word and create a new document.

            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
        }
        void EnterPoints()
        {
            while (true)
            {
                List<Point> temp = new List<Point>();
                Console.WriteLine("Введите группу точек? 111 - нет, 1 - да");
                int digit = int.Parse(Console.ReadLine());
                if (digit == 111) break;
                while (true)
                {
                    Console.WriteLine("Введите точку, когда закончите введите в у 111:");
                    Console.Write("X = ");
                    int x = int.Parse(Console.ReadLine());
                    Console.Write("Y = ");
                    int y = int.Parse(Console.ReadLine());

                    if (y == 111) break;
                    temp.Add(new Point(x, y));
                }
                if (temp.Count != 0)
                {
                    points.Add(new List<Point>(temp));
                }
                temp.Clear();
            }
        }
        void AddTextToWord(Word.Paragraph para, string text)
        {
            para.Range.Text = text;
            para.Range.InsertParagraphAfter();
        }

        void SaveDoc()
        {
            object filename = Path.GetFullPath("examtask.doc");
            oDoc.SaveAs(ref filename, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing);
            Console.WriteLine("Ваш документ готов! Он находится в bin->Debug");
        }

        public void FirstSolutionFunction()
        {
            EnterPoints();
            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            AddTextToWord(para, "Решающие функции первого вида");
            AddTextToWord(para, "x - x1 / x2 - x1 = y - y1 / y2 - y1");



            SaveDoc();
        }
    }
}
