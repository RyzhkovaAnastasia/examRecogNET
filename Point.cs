namespace ExamRecog
{
    internal class Point
    {
        public int PointClass { get; set; }

        public double X { get; set; }

        public double Y { get; set; }

        public Point(double x, double y,int pointClass)
        {
            X = x;
            Y = y;
            PointClass = pointClass;
        }
    }
}