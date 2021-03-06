using System;

namespace ExamRecog
{
    class Program
    {
        static void Main(string[] args)
        {
            Solver solver = new Solver();
            int number = 0;
            Console.WriteLine("Введите номер задачи:");
            Console.WriteLine("1. Решающие фукции 1 вид:");
            Console.WriteLine("2. Решающие функции 2 вид:");
            Console.WriteLine("3. Решающие функции 3 вид:");
            Console.WriteLine("4. Кластеризация пороговым методом:");
            Console.WriteLine("5. Кластеризация maxmin:");
            Console.WriteLine("6. Кластеризация к-внутригрупповых средних:");
            Console.WriteLine("7. Обучение перцептрона с дробной коррекцией весов:");
            Console.WriteLine("8. Обучение перцептрона с фиксированным приращением:");
            Console.WriteLine("9. Обучение перцептрона, алгоритм коррекции абсолютной величины:");
            number = int.Parse(Console.ReadLine());
            solver.EnterPoints();
            switch (number)
            {
                case 1:
                    solver.FirstSolutionFunction();
                    break;
                case 2:
                    break;
                case 3:
                    solver.ThirdSolutionFunction();
                    break;
                case 4:
                    Console.WriteLine("Введите пороговое значение/n (диапазон от максимального внутрикластерного до минимального межкластерного, дробь через запятую).");
                    var threshold = double.Parse(Console.ReadLine());
                    solver.SimpleThreshold(threshold);
                    break;
                case 5:
                    solver.Maximin();
                    break;
                case 6:
                    solver.KIntergroupAverage();
                    break;
                case 7:
                    solver.PerceptronFrac();
                    break;
                case 8:
                    solver.PerceptronFix();
                    break;
                case 9:
                    solver.PerceptronWeight();
                    break;

            }
            Console.Read();
        }
    }
}
