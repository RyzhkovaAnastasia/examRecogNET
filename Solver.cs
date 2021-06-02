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
            Console.WriteLine("Дроби вводятся через точку!");
            Console.WriteLine("Порядок точек может иметь значение, вводите, начиная с левого верхнего угла (верхние точки имеют приоритет).");
            do
            {
                while (true)
                {
                    Console.WriteLine("Введите точку, когда закончите введите в у 111:");
                    Console.Write("X = ");
                    double x = double.Parse(Console.ReadLine());
                    Console.Write("Y = ");
                    double y = double.Parse(Console.ReadLine());

                    if (y == 111.0) break;
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
            oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            oDoc = null;
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            oWord = null;
        }

        public void FirstSolutionFunction()
        {
            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            AddTextToWord(para, "Решающие функции первого вида");
            AddTextToWord(para, "x - x1 / x2 - x1 = y - y1 / y2 - y1");



            SaveDoc();
        }

        public void ThirdSolutionFunction()
        {
            int answ = 0;
            int @class = 0;
            bool flag = true;
            List<Point> etalon = new List<Point>();
            do
            {
                Console.WriteLine("Помощь в выборе эталонов. Введите предполагаемые эталоны:");
                while (true)
                {
                    Console.WriteLine("Введите точку, когда закончите введите в у 111:");
                    Console.Write("X = ");
                    double x = double.Parse(Console.ReadLine());
                    Console.Write("Y = ");
                    double y = double.Parse(Console.ReadLine());

                    if (y == 111) break;
                    etalon.Add(new Point(x, y, @class++));
                }

                for (int i = 0; i < points.Count; i++)
                {
                    for (int j = 0; j < etalon.Count; j++)
                    {
                        if (points[i].PointClass == etalon[j].PointClass)
                        {
                            double t = etalon[j].X * points[i].X + etalon[j].Y * points[i].Y - (0.5 * (etalon[j].X * etalon[j].X + etalon[j].Y * etalon[j].Y));
                            if (t < 0)
                            {
                                Console.WriteLine($"X{i + 1} class = {points[i].PointClass + 1} for etalon class {etalon[j].PointClass}=> {t} - неудача");
                                flag = false;
                            }
                            else
                            {
                                Console.WriteLine($"X{i + 1} class = {points[i].PointClass + 1} for etalon class {etalon[j].PointClass}=> {t}");
                            }
                        }
                        else if (points[i].PointClass != etalon[j].PointClass)
                        {
                            double t = (etalon[j].X * points[i].X + etalon[j].Y * points[i].Y) - (0.5 * (etalon[j].X * etalon[j].X + etalon[j].Y * etalon[j].Y));
                            if (t > 0)
                            {
                                Console.WriteLine($"X{i + 1} class = {points[i].PointClass + 1} for etalon class {etalon[j].PointClass}=> {t} - неудача");
                                flag = false;
                            }
                            else
                            {
                                Console.WriteLine($"X{i + 1} class = {points[i].PointClass + 1} for etalon class {etalon[j].PointClass}=> {t}");
                            }
                        }
                    }
                }
                Console.WriteLine("Повторить? да - 1, нет - 111");
                answ = int.Parse(Console.ReadLine());

            } while (answ != 111);
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
                if (min > maximin && !clusterCenters.Contains(i))
                {
                    maximin = min;
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

            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            para.Range.Text += "Метод k-внутригрупповых средних.";
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
                        distances[i][j] = CalculateDistances(points[i], clusterCenters[j]);
                    }
                }
                ClassesAllocation(clusterCenters);
                var temp = "min (";
                for (int i = 0; i < points.Count; i++)
                {
                    for (int j = 0; j < clusterCenters.Count; j++)
                    {
                        temp += "ρ(X" + j + "X" + i + ");";
                    }
                    para.Range.Text += temp + ") ="+ "ρ(X" +points[i].PointClass+"X"+i+")"+ "\nЗначит X" + i + "є ω" + points[i].PointClass;
                    temp = "min (";
                }

                for (int i = 0; i < clusterCenters.Count; i++)
                {
                    clusterCenters[i].X = 0;
                    clusterCenters[i].Y = 0;
                }
                para.Range.Text += "Пересчитываем центры кластеров.";
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
                    para.Range.Text += "Z" + i + "," + 1 + " = " + clusterCenters[i].X;
                    para.Range.Text += "Z" + i + "," + 2 + " = " + clusterCenters[i].Y;
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
            var textBuff = string.Empty;
            para.Range.Text += "Результат:";
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
            para.Range.Text += "Метод максимин.";
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
                var iter = 0;
                for (int i = 0; i < clusterCenters.Count; i++)
                {
                    for (int j = i + 1; j < clusterCenters.Count; j++)
                    {
                        centresDistancesSum += distances[clusterCenters[i]][clusterCenters[j]];
                        iter++;
                    }
                }
                centresDistancesSum = centresDistancesSum / 2 / iter;

                para.Range.Text += "Ищем максимум от найденных минимумов, он соответствует расстоянию между " + minMaxInd1 + " и " + minMaxInd2 + " и равен " + distances[minMaxInd1][minMaxInd2];
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

        public void SimpleThreshold(double threshold)
        {
            var pointClass = 0;
            var clusterCenters = new List<int>() { 0 };
            var currentDistance = 0.0;
            var clusterCenterAppeared = true;
            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            para.Range.Text += "Метод простого порогового значения.\nРассматриваем точки сверху вниз, слева направо.";
            para.Range.Text += "Выбираем порог в диапазоне [;] - T = " + threshold;
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = 0; j < clusterCenters.Count; j++)
                {
                    currentDistance = CalculateDistances(points[i], points[clusterCenters[j]]);
                    if (currentDistance < threshold)
                    {
                        points[i].PointClass = points[clusterCenters[j]].PointClass;
                        clusterCenterAppeared = false;
                    }
                }
                if (clusterCenterAppeared && !clusterCenters.Contains(i))
                {
                    points[i].PointClass = ++pointClass;
                    clusterCenters.Add(i);
                    para.Range.Text += "Точка X" + i + " - центр нового кластера.";
                    i = -1;
                }
                else
                    para.Range.Text += "Точка X" + i + "є ω" + points[i].PointClass;
                clusterCenterAppeared = true;
            }
            var textBuff = string.Empty;
            para.Range.Text += "Результат:";
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

        public void PerceptronFix()
        {
            Console.WriteLine("НЕ ОТПРАВЛЯЙТЕ ПОЛНОЕ РЕШЕНИЕ, на 3-4 итерации закончите файл. Конец файла должен содержать слова о условии конца итераций:");
            Console.WriteLine("Повторяем итерации до тех пор, пока перцептрон будет поощряться и наказываться. Если в итерации перцептрон изменен не был, это означает конец решения.");
            double[] W = new double[] { 1, 1, 1 };

            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            AddTextToWord(para, "Обучение перцепрона с фиксированным приращением, C = 1");
            for (int i = 0; i < points.Count; i++)
            {
                AddTextToWord(para, "X" + (i + 1) + " = [" + points[i].X + ";" + points[i].Y + ";1] принадлежит классу №" + (points[i].PointClass + 1));
            }
            AddTextToWord(para, "d1(X1)=W1`*X W1`(0)=[1;1;1]");


            int step = 1;
            int @class = 0;

            for (@class = 0; @class <= points[points.Count - 1].PointClass; @class++)
            {
                AddTextToWord(para, $"Найдем решающую функцию #{@class + 1} \n");
                bool flag = true; //флаг завершения алгоритма
                W[0] = 1;
                W[1] = 1;
                W[2] = 1;
                step = 1;

                for (int e = 0; e < 5 && flag; e++)
                { // пока вектор W будет изменяться
                    flag = false;
                    for (int i = 0; i < points.Count; i++)
                    {
                        double res = points[i].X * W[0] + points[i].Y * W[1] + 1 * W[2];
                        AddTextToWord(para, $"d{@class + 1}(X{i + 1})= [{W[0]};{W[1]};{W[2]}] * [{points[i].X};{points[i].Y};1]^-1 = {res}");

                        if (points[i].PointClass == @class && res > 0)
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} > 0, поэтому перцептрон оставляем без изменений.W{@class + 1}({step}) = W{@class + 1}({step - 1}) \n");
                        }
                        else if (points[i].PointClass == @class && res <= 0)
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} < 0, должно быть > 0, поэтому перцептрон поощряем. W{@class + 1}({step}) = W{@class + 1}({step - 1}) + X{i + 1} " +
                               $"= [{W[0]};{W[1]};{W[2]}]^-1 + [{points[i].X};{points[i].Y};1]^-1 = [{W[0] + points[i].X};{W[1] + points[i].Y};{W[2] + 1}]^-1 \n");
                            W[0] = W[0] + points[i].X;
                            W[1] = W[1] + points[i].Y;
                            W[2] = W[2] + 1;
                            flag = true;
                        }
                        else if (points[i].PointClass != @class && res >= 0)
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} > 0, должно быть < 0, поэтому перцептрон наказываем. W{@class + 1}({step}) = W{@class + 1}({step - 1}) - X{i + 1} " +
                                $"= [{W[0]};{W[1]};{W[2]}]^-1 - [{points[i].X};{points[i].Y};1]^-1 = [{W[0] - points[i].X};{W[1] - points[i].Y};{W[2] - 1}]^-1 \n");
                            W[0] = W[0] - points[i].X;
                            W[1] = W[1] - points[i].Y;
                            W[2] = W[2] - 1;
                            flag = true;
                        }
                        else
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} < 0, поэтому перцептрон оставляем без изменений.W1({step}) = W1({step - 1}) \n");
                        }
                        step++;
                    }

                }
                if (flag) AddTextToWord(para, $"И так далее пока прохождение всех точек не приведет ни к наказанию, ни к поощрению перцептрона. \n");
            }
            SaveDoc();
        }


        public void PerceptronWeight()
        {
            Console.WriteLine("НЕ ОТПРАВЛЯЙТЕ ПОЛНОЕ РЕШЕНИЕ, на 3-4 итерации закончите файл. Конец файла должен содержать слова о условии конца итераций:");
            Console.WriteLine("Повторяем итерации до тех пор, пока перцептрон будет поощряться и наказываться. Если в итерации перцептрон изменен не был, это означает конец решения.");
            double[] W = new double[] { 1, 1, 1 };
            double C = 0; // приращение

            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            AddTextToWord(para, "Обучение перцепрона с коррекцией весов");
            for (int i = 0; i < points.Count; i++)
            {
                AddTextToWord(para, "X" + (i + 1) + " = [" + points[i].X + ";" + points[i].Y + ";1] принадлежит классу №" + (points[i].PointClass + 1));
            }
            AddTextToWord(para, "d1(X1)=W1`*X W1`(0)=[1;1;1]");


            int step = 1;
            int @class = 0;

            for (@class = 0; @class <= points[points.Count - 1].PointClass; @class++)
            {
                AddTextToWord(para, $"Найдем решающую функцию #{@class + 1} \n");
                bool flag = true; //флаг завершения алгоритма
                W[0] = 1;
                W[1] = 1;
                W[2] = 1;
                step = 1;

                for (int e = 0; e < 5 && flag; e++)
                { // пока вектор W будет изменяться
                    flag = false;
                    for (int i = 0; i < points.Count; i++)
                    {
                        double res = points[i].X * W[0] + points[i].Y * W[1] + 1 * W[2];
                        AddTextToWord(para, $"d{@class + 1}(X{i + 1})= [{W[0]};{W[1]};{W[2]}] * [{points[i].X};{points[i].Y};1]^-1 = {res}");

                        if (points[i].PointClass == @class && res > 0)
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} > 0, поэтому перцептрон оставляем без изменений.W{@class + 1}({step}) = W{@class + 1}({step - 1}) \n");
                        }
                        else if (points[i].PointClass == @class && res <= 0)
                        {
                            C = Math.Ceiling(Convert.ToDouble(Math.Abs(res) / (points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1)));
                            if (C == 0) C = 1;
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} < 0, должно быть > 0, поэтому перцептрон поощряем. W{@class + 1}({step}) = W{@class + 1}({step - 1}) + C*X{i + 1}\n " +
                               $"Вычисляем С: С = W{@class + 1}*X{i + 1} / (X{i + 1} * X`{i + 1}) ={res} / {(points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1)} = {C}\n" +
                                $"[{W[0]};{W[1]};{W[2]}]^-1 + {C}*[{points[i].X};{points[i].Y};1]^-1 = [{W[0] + C * points[i].X};{W[1] + C * points[i].Y};{W[2] + C * 1}]^-1 \n"); ;
                            W[0] = W[0] + C * points[i].X;
                            W[1] = W[1] + C * points[i].Y;
                            W[2] = W[2] + C * 1;
                            flag = true;
                        }
                        else if (points[i].PointClass != @class && res >= 0)
                        {
                            C = Math.Ceiling(Convert.ToDouble(Math.Abs(res) / (points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1)));
                            if (C == 0) C = 1;
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} > 0, должно быть < 0, поэтому перцептрон наказываем. W{@class + 1}({step}) = W{@class + 1}({step - 1}) - C*X{i + 1} \n" +
                                $"Вычисляем С: С = W{@class + 1}*X{i + 1} / (X{i + 1} * X`{i + 1}) ={res} / {(points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1)} = {C}\n" +
                                $"[{W[0]};{W[1]};{W[2]}]^-1 - {C}*[{points[i].X};{points[i].Y};1]^-1 = [{W[0] - C * points[i].X};{W[1] - C * points[i].Y};{W[2] - C * 1}]^-1 \n");
                            W[0] = W[0] - C * points[i].X;
                            W[1] = W[1] - C * points[i].Y;
                            W[2] = W[2] - C * 1;
                            flag = true;
                        }
                        else
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} < 0, поэтому перцептрон оставляем без изменений.W1({step}) = W1({step - 1}) \n");
                        }
                        step++;
                    }

                }
                if (flag) AddTextToWord(para, $"И так далее пока прохождение всех точек не приведет ни к наказанию, ни к поощрению перцептрона. \n");
            }

            SaveDoc();

        }

        public void PerceptronFrac()
        {
            Console.WriteLine("НЕ ОТПРАВЛЯЙТЕ ПОЛНОЕ РЕШЕНИЕ, на 3-4 итерации закончите файл. Конец файла должен содержать слова о условии конца итераций:");
            Console.WriteLine("Повторяем итерации до тех пор, пока перцептрон будет поощряться и наказываться. Если в итерации перцептрон изменен не был, это означает конец решения.");
            double[] W = new double[] { 1, 1, 1 };
            double C = 0; // приращение

            Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
            AddTextToWord(para, "Обучение перцепрона с коррекцией весов");
            for (int i = 0; i < points.Count; i++)
            {
                AddTextToWord(para, "X" + (i + 1) + " = [" + points[i].X + ";" + points[i].Y + ";1] принадлежит классу №" + (points[i].PointClass + 1));
            }
            AddTextToWord(para, "d1(X1)=W1`*X W1`(0)=[1;1;1]");


            int step = 1;
            int @class = 0;

            for (@class = 0; @class <= points[points.Count - 1].PointClass; @class++)
            {
                AddTextToWord(para, $"Найдем решающую функцию #{@class + 1} \n");
                bool flag = true; //флаг завершения алгоритма
                W[0] = 1;
                W[1] = 1;
                W[2] = 1;
                step = 1;

                for (int e = 0; e < 5 && flag; e++)
                { // пока вектор W будет изменяться
                    flag = false;
                    for (int i = 0; i < points.Count; i++)
                    {
                        double res = points[i].X * W[0] + points[i].Y * W[1] + 1 * W[2];
                        AddTextToWord(para, $"d{@class + 1}(X{i + 1})= [{W[0]};{W[1]};{W[2]}] * [{points[i].X};{points[i].Y};1]^-1 = {res}");

                        if (points[i].PointClass == @class && res > 0)
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} > 0, поэтому перцептрон оставляем без изменений.W{@class + 1}({step}) = W{@class + 1}({step - 1}) \n");
                        }
                        else if (points[i].PointClass == @class && res <= 0)
                        {
                            C = Math.Round(Math.Abs(res) / (points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1), 2);
                            if (C == 0) C = 1;
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} < 0, должно быть > 0, поэтому перцептрон поощряем. W{@class + 1}({step}) = W{@class + 1}({step - 1}) + C*X{i + 1}\n " +
                               $"Вычисляем С: С = W{@class + 1}*X{i + 1} / (X{i + 1} * X`{i + 1}) ={res} / {(points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1)} = {C}\n" +
                                $"[{W[0]};{W[1]};{W[2]}]^-1 + {C}*[{points[i].X};{points[i].Y};1]^-1 = [{Math.Round(W[0] + C * points[i].X),2};{Math.Round(W[1] + C * points[i].Y, 2)};{Math.Round(W[2] + C * 1, 2)}]^-1 \n"); ;
                            W[0] = Math.Round(W[0] + C * points[i].X, 2);
                            W[1] = Math.Round(W[1] + C * points[i].Y, 2);
                            W[2] = Math.Round(W[2] + C * 1, 2);
                            flag = true;
                        }
                        else if (points[i].PointClass != @class && res >= 0)
                        {
                            C = Math.Round(Math.Abs(res) / (points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1), 2);
                            if (C == 0) C = 1;
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} > 0, должно быть < 0, поэтому перцептрон наказываем. W{@class + 1}({step}) = W{@class + 1}({step - 1}) - C*X{i + 1} \n" +
                                $"Вычисляем С: С = W{@class + 1}*X{i + 1} / (X{i + 1} * X`{i + 1}) ={res} / {(points[i].X * points[i].X + points[i].Y * points[i].Y + 1 * 1)} = {C}\n" +
                                $"[{W[0]};{W[1]};{W[2]}]^-1 - {C}*[{points[i].X};{points[i].Y};1]^-1 = [{Math.Round(W[0] - C * points[i].X, 2)};{Math.Round(W[1] - C * points[i].Y, 2)};{Math.Round(W[2] - C * 1, 2)}]^-1 \n");
                            W[0] = Math.Round(W[0] - C * points[i].X, 2);
                            W[1] = Math.Round(W[1] - C * points[i].Y, 2);
                            W[2] = Math.Round(W[2] - C * 1, 2);
                            flag = true;
                        }
                        else
                        {
                            AddTextToWord(para, $"d{@class + 1}(X{i + 1})={res} < 0, поэтому перцептрон оставляем без изменений.W1({step}) = W1({step - 1}) \n");
                        }
                        step++;
                    }

                }
                if (flag) AddTextToWord(para, $"И так далее пока прохождение всех точек не приведет ни к наказанию, ни к поощрению перцептрона. \n");
            }

            SaveDoc();

        }
    }
}
