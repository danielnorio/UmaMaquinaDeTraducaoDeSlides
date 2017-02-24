namespace TranslationMachine
{
    class Rectangle
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }

        public Rectangle (float _Left, float _Top, float _Width, float _Height)
        {
            Left = _Left;
            Top = _Top;
            Width = _Width;
            Height = _Height;
        }
    }
}
