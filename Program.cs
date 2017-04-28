using System;
using System.IO;
using CLP = CommandLineParser;
using Microsoft.Extensions.Configuration;

namespace exchange_flagged_histogram
{
    class Program
    {
        static void Main(string[] args)
        {
            var config = new CLP.Arguments.FileArgument('c', "config")
            {
                DefaultValue = new FileInfo("config.json")
            };

            var commandLineParser = new CLP.CommandLineParser()
            {
                Arguments = {
                    config,
                }
            };

            try
            {
                commandLineParser.ParseCommandLine(args);
                commandLineParser.ShowParsedArguments();

                var configuration = new ConfigurationBuilder()
                    .AddJsonFile(config.Value.FullName, true)
                    .Build();
            }
            catch (CLP.Exceptions.CommandLineException e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
