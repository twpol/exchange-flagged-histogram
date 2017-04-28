using System;
using System.IO;
using CLP = CommandLineParser;
using Microsoft.Exchange.WebServices.Data;
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

                Main(new ConfigurationBuilder()
                    .AddJsonFile(config.Value.FullName, true)
                    .Build());
            }
            catch (CLP.Exceptions.CommandLineException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        static void Main(IConfigurationRoot config)
        {
            var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            var credentials = config.GetSection("credentials");
            service.Credentials = new WebCredentials(credentials["username"], credentials["password"]);
            service.AutodiscoverUrl(credentials["email"], ValidateHTTPSUri);
        }

        static bool ValidateHTTPSUri(string redirectionUri)
        {
            return new Uri(redirectionUri).Scheme == "https";
        }
    }
}
