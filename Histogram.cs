using System;
using System.Collections.Generic;
using System.Linq;

namespace exchange_flagged_histogram
{
    class Histogram
    {
        List<char> Categories;
        Dictionary<char, List<double>> Values;

        public Histogram(List<char> categories)
        {
            Categories = categories;
            Values = new Dictionary<char, List<double>>();
            foreach (var category in Categories)
                Values[category] = new List<double>();
        }

        public void Add(char category, double value)
        {
            if (Values.TryGetValue(category, out var values))
                values.Add(value);
        }

        public void RenderTo(HistogramOutput output, List<char> valueCategories)
        {
            var valuesByIndex = Categories
                .Select(category => Values[category])
                .ToList();

            var valueCategoryIndexes = valueCategories
                .Select(valueCategory => Categories
                    .FindIndex(category => category == valueCategory));

            var minimum = double.MaxValue;
            var maximum = double.MinValue;
            foreach (var values in Values.Values)
            {
                values.Sort();
                if (values.Count > 0)
                {
                    minimum = Math.Min(minimum, values[0]);
                    maximum = Math.Max(maximum, values[values.Count - 1]);
                }
            }

            if (output.BinSize > 0)
            {
                output.Base = (int)Math.Floor(minimum / output.BinSize) * output.BinSize;
                output.Height = (int)Math.Ceiling((maximum - output.Base) / output.BinSize);
            }
            else if (output.Height > 0)
            {
                output.Base = (int)Math.Floor(minimum);
                output.BinSize = (int)Math.Ceiling((maximum - output.Base) / output.Height);
            }
            else
            {
                throw new InvalidOperationException("Histogram.Render must have either Bin or Height set to a positive integer.");
            }

            var maximumCount = 0;
            var counts = new int[output.Height, Categories.Count];
            var totalCounts = new int[output.Height];
            for (var categoryIndex = 0; categoryIndex < Categories.Count; categoryIndex++)
            {
                for (var valueIndex = 0; valueIndex < valuesByIndex[categoryIndex].Count; valueIndex++)
                {
                    var bin = (int)Math.Floor((valuesByIndex[categoryIndex][valueIndex] - output.Base) / output.BinSize);
                    counts[bin, categoryIndex]++;
                    totalCounts[bin]++;
                    maximumCount = Math.Max(maximumCount, totalCounts[bin]);
                }
            }

            output.Scale = Math.Max(Math.Min((double)output.Width / maximumCount, output.MaxScale), output.MinScale);
            output.Values = new int[output.Height];
            output.Graph = new string[output.Height];

            for (var line = 0; line < output.Graph.Length; line++)
            {
                output.Values[line] = 0;
                foreach (var categoryIndex in valueCategoryIndexes)
                {
                    output.Values[line] += counts[line, categoryIndex];
                }
                var graphError = 0D;
                output.Graph[line] = "";
                for (var categoryIndex = 0; categoryIndex < Categories.Count; categoryIndex++)
                {
                    var size = graphError + (double)counts[line, categoryIndex] * output.Scale;
                    var sizeRound = (int)Math.Round(size);
                    graphError = size - sizeRound;
                    output.Graph[line] += new String(Categories[categoryIndex], sizeRound);
                }
            }
        }
    }

    class HistogramOutput
    {
        public int Base;
        public int BinSize;
        public double MinScale = double.MinValue;
        public double MaxScale = double.MaxValue;
        public double Scale;
        public int Width;
        public int Height;
        public int[] Values;
        public string[] Graph;
    }
}
