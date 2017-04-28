using System;
using System.Collections.Generic;

namespace exchange_flagged_histogram
{
    class Histogram
    {
        char[] Categories;
        Dictionary<char, List<double>> Values;

        public Histogram(char[] categories)
        {
            Categories = categories;
            Values = new Dictionary<char, List<double>>();
        }

        public void Add(char category, double value)
        {
            if (Values.TryGetValue(category, out var values))
                values.Add(value);
        }
    }
}
