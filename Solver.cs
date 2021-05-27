using System;
using System.Collections.Generic;
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
        List<Point> points = new List<Point>();
        List<List<double>> distances = new List<List<double>>(); //first element contains distances of first point to each other etc
        public Solver()
        {
            //Start Word and create a new document.

            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
        }
        public void EnterPoints()
        {
            var pointClass = 0;

            do
            {
                while (true)
                {
                    Console.WriteLine("Введите точку, когда закончите введите в у 111:");
                    Console.Write("X = ");
                    int x = int.Parse(Console.ReadLine());
                    Console.Write("Y = ");
                    int y = int.Parse(Console.ReadLine());

                    if (y == 111) break;
                    points.Add(new Point(x, y, pointClass));
                }
                pointClass++;
                foreach (Point p in points)
                    distances.Add(new List<double>());
                Console.WriteLine("Ввести новую группу точек? 111 - нет, 1 - да");
                int digit = int.Parse(Console.ReadLine());
                if (digit == 111) break;
            } while (true);
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
            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            AddTextToWord(para, "Решающие функции первого вида");
            AddTextToWord(para, "x - x1 / x2 - x1 = y - y1 / y2 - y1");



            SaveDoc();
        }

        public void CalculateDistances()
        {
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < points.Count; j++)
                {
                    distances[i].Add(Math.Sqrt(Math.Pow(points[i].X - points[j].X, 2) + Math.Pow(points[i].Y - points[j].Y, 2)));
                }
            }
        }

        private void ClassesAllocation(out int ind1, out int ind2, List<int> clusterCentres)
        {
            double min = 100.0, maximin = 0.0;
            int minInd1 = -1, minInd2 = -1;
            ind1 = 0;
            ind2 = 0;
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < clusterCentres.Count - 1; j++)
                {
                    if (!clusterCentres.Contains(i))
                    {
                        if (distances[i][clusterCentres[j]] >= distances[i][clusterCentres[j + 1]])
                        {
                            if (distances[i][clusterCentres[j + 1]] < min)
                            {
                                points[i].PointClass = points[clusterCentres[j + 1]].PointClass;
                                min = distances[i][clusterCentres[j + 1]];
                                minInd1 = i;
                                minInd2 = clusterCentres[j + 1];
                            }
                        }
                        else
                        {

                            if (distances[i][clusterCentres[j]] < min)
                            {
                                points[i].PointClass = points[clusterCentres[j]].PointClass;
                                min = distances[i][clusterCentres[j]];
                                minInd1 = i;
                                minInd2 = clusterCentres[j];
                            }
                        }
                    }
                }
                if (min > maximin)
                {
                    ind1 = minInd1;
                    ind2 = minInd2;
                }
                min = 100;
            }
        }
        public void Maximin()
        {
            var max = 0.0;
            int maxInd1 = -1, maxInd2 = -1, minMaxInd1, minMaxInd2;
            var clusterCentres = new List<int>();
            CalculateDistances();

            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < points.Count; j++)
                {
                    if (distances[i][j] > max)
                    {
                        max = distances[i][j];
                        maxInd1 = i;
                        maxInd2 = j;
                    }
                }
            }
            max = 0.0;
            clusterCentres.Add(maxInd1);
            clusterCentres.Add(maxInd2);
            points[maxInd1].PointClass = 0;
            points[maxInd2].PointClass = 1;
            do
            {
                ClassesAllocation(out minMaxInd1, out minMaxInd2, clusterCentres);

                var centresDistancesSum = 0.0;
                for (int i = 0; i < clusterCentres.Count; i++)
                {
                    for (int j = 0; j < clusterCentres.Count; j++)
                    {
                        centresDistancesSum += distances[clusterCentres[i]][clusterCentres[j]];
                    }
                }
                centresDistancesSum = centresDistancesSum / clusterCentres.Count / Math.Pow(2, clusterCentres.Count - 1);
                if (distances[minMaxInd1][minMaxInd2] > centresDistancesSum)
                {
                    clusterCentres.Add(minMaxInd1);
                    points[minMaxInd1].PointClass = clusterCentres.Count - 1;
                }
                else break;
            } while (true);
            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            for (int i = 0; i < clusterCentres.Count; i++)
            {
                para.Range.Text+= "ω_" + i + "= {";
                for (int j = 0; j < points.Count; j++)
                {
                    if (points[j].PointClass == i)
                        para.Range.Text+= "X_" + j + " ";
                }
                para.Range.Text+= "} ";
            }
            SaveDoc();
        }
    }
}
