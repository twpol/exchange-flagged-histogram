using System;
using System.Collections.Generic;

namespace exchange_flagged_histogram
{
    class Histogram
    {
        char[] Categories;
        List<double>[] Values;

        public Histogram(char[] categories)
        {
            Categories = categories;
            Values = new List<double>[categories.Length];
            for (var i = 0; i < Values.Length; i++)
                Values[i] = new List<double>();
        }

        public void Add(int category, double value)
        {
            Values[category].Add(value);
        }
    }
}
