namespace ExamRecog
{
    internal class Point
    {
        public int PointClass { get; set; }
        public int X { get; private set; }
        public int Y { get; private set; }

        public Point(int x, int y,int pointClass)
        {
            X = x;
            Y = y;
            PointClass = pointClass;
        }
    }
}