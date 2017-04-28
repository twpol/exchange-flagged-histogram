using System;

namespace exchange_flagged_histogram
{
    class Program
    {
        static void Main(string[] args)
        {
            var commandLineParser = new CommandLineParser.CommandLineParser()
            {
                Arguments = {
                }
            };
            try
            {
                commandLineParser.ParseCommandLine(args);
                commandLineParser.ShowParsedArguments();
            }
            catch (CommandLineParser.Exceptions.CommandLineException e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
