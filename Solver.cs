﻿using System;
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

        public double CalculateDistances(Point p1, Point p2)
        {
            return Math.Sqrt(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2));
        }

        private void ClassesAllocation(out int ind1, out int ind2, List<int> clusterCenters)
        {
            double min = 100.0, maximin = 0.0;
            int minInd1 = -1, minInd2 = -1;
            ind1 = 0;
            ind2 = 0;
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < clusterCenters.Count; j++)
                {
                    if (!clusterCenters.Contains(i))
                    {
                        if (distances[i][clusterCenters[j]] < min)
                        {
                            points[i].PointClass = points[clusterCenters[j]].PointClass;
                            min = distances[i][clusterCenters[j]];
                            minInd1 = i;
                            minInd2 = clusterCenters[j];
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

        private void ClassesAllocation(List<Point> clusterCenters)
        {
            double min = 100.0;
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < clusterCenters.Count; j++)
                {
                    if (!clusterCenters.Contains(points[i]))
                    {
                        if (distances[i][j] < min)
                        {
                            points[i].PointClass = clusterCenters[j].PointClass;
                            min = distances[i][j];
                        }
                    }
                }
                min = 100;
            }
        }

        public void KIntergroupAverage()
        {
            bool changed;
            var iter = 0;
            var clusterCenters = new List<Point>() { new Point(points[0].X, points[0].Y, points[0].PointClass),
                new Point(points[1].X, points[1].Y, points[1].PointClass), new Point(points[2].X, points[2].Y, points[2].PointClass) };
            clusterCenters[0].PointClass = 0;
            clusterCenters[1].PointClass = 1;
            clusterCenters[2].PointClass = 2;
            var clusterCentersOld = new List<Point>() { new Point(clusterCenters[0].X, clusterCenters[0].Y, clusterCenters[0].PointClass),
                new Point(clusterCenters[1].X, clusterCenters[1].Y, clusterCenters[1].PointClass), new Point(clusterCenters[2].X, clusterCenters[2].Y, clusterCenters[2].PointClass) };
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < clusterCenters.Count; j++)
                {
                    distances[i].Add(0);
                }
            }
            do
            {
                changed = false;
                for (int i = 0; i < points.Count; i++)
                {
                    for (int j = 0; j < clusterCenters.Count; j++)
                    {
                        distances[i][j]=CalculateDistances(points[i], clusterCenters[j]);
                    }
                }
                ClassesAllocation(clusterCenters);
                for (int i = 0; i < clusterCenters.Count; i++)
                {
                    clusterCenters[i].X = 0;
                    clusterCenters[i].Y = 0;
                }
                for (int i = 0; i < clusterCenters.Count; i++)
                {
                    for (int j = 0; j < points.Count; j++)
                    {
                        if (points[j].PointClass == i)
                        {
                            iter++;
                            clusterCenters[i].X += points[j].X;
                            clusterCenters[i].Y += points[j].Y;
                        }
                    }
                    clusterCenters[i].X = clusterCenters[i].X / iter;
                    clusterCenters[i].Y = clusterCenters[i].Y / iter;
                    iter = 0;
                }
                for (int i = 0; i < clusterCenters.Count; i++)
                {
                    if (clusterCenters[i].X == clusterCentersOld[i].X && clusterCenters[i].Y == clusterCentersOld[i].Y)
                        continue;
                    else
                    {
                        changed = true;
                        clusterCentersOld[i].X = clusterCenters[i].X;
                        clusterCentersOld[i].Y = clusterCenters[i].Y;
                    }
                }

            } while (changed);

        }

        public void Maximin()
        {
            var max = 0.0;
            int maxInd1 = -1, maxInd2 = -1, minMaxInd1, minMaxInd2;
            var clusterCenters = new List<int>();
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < points.Count; j++)
                {
                    distances[i].Add(CalculateDistances(points[i], points[j]));
                }
            }

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
            clusterCenters.Add(maxInd1);
            clusterCenters.Add(maxInd2);
            points[maxInd1].PointClass = 0;
            points[maxInd2].PointClass = 1;
            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            var textBuff = string.Empty;
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < points.Count; j++)
                {
                    para.Range.Text += "ρ(X" + i + " X" + j + ") = " + distances[i][j];
                }
                para.Range.Text += "\n";
            }
            do
            {
                ClassesAllocation(out minMaxInd1, out minMaxInd2, clusterCenters);

                var centresDistancesSum = 0.0;
                for (int i = 0; i < clusterCenters.Count; i++)
                {
                    for (int j = 0; j < clusterCenters.Count; j++)
                    {
                        centresDistancesSum += distances[clusterCenters[i]][clusterCenters[j]];
                    }
                }
                centresDistancesSum = centresDistancesSum / clusterCenters.Count / Math.Pow(2, clusterCenters.Count - 1);

                para.Range.Text += "Ищем максимум от найденных минимумов, он соответствует " + minMaxInd1 + " и равен " + distances[minMaxInd1][minMaxInd2];
                textBuff += "Половина среднего расстояния между известными центрами кластеров равна " + centresDistancesSum + ". Поскольку " + distances[minMaxInd1][minMaxInd2];
                if (distances[minMaxInd1][minMaxInd2] > centresDistancesSum)
                {
                    clusterCenters.Add(minMaxInd1);
                    points[minMaxInd1].PointClass = clusterCenters.Count - 1;
                    para.Range.Text += textBuff + ">" + centresDistancesSum + ", то X" + minMaxInd1 + " - новый центр кластеров.";
                }
                else
                {
                    para.Range.Text += textBuff + "<" + centresDistancesSum + ", то новых кластеров нет.";
                    break;
                }
                textBuff = string.Empty;
            } while (true);
            textBuff = string.Empty;
            for (int i = 0; i < clusterCenters.Count; i++)
            {
                for (int j = 0; j < points.Count; j++)
                {
                    if (points[j].PointClass == i)
                        textBuff += "X_" + j + " ";
                }

                para.Range.Text += "ω_" + i + "= {" + textBuff + "} ";
                textBuff = string.Empty;
            }
            SaveDoc();
        }
    }
}