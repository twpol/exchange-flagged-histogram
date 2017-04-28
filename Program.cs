using System;
using System.Collections.Generic;
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
            var service = new ExchangeService(ExchangeVersion.Exchange2013);

            LogIn(config.GetSection("credentials"), service);

            Histogram(config.GetSection("histogram"), service);
        }

        private static void LogIn(IConfigurationSection config, ExchangeService service)
        {
            service.Credentials = new WebCredentials(config["username"], config["password"]);
            service.AutodiscoverUrl(config["email"], redirectionUri =>
                new Uri(redirectionUri).Scheme == "https"
            );
        }

        private static void Histogram(IConfigurationSection config, ExchangeService service)
        {
            var categories = new List<char>(4);
            if ((config["includeFlaggedOld"] ?? "true") == "true")
                categories.Add('#');
            if ((config["includeFlaggedNew"] ?? "true") == "true")
                categories.Add('+');
            if ((config["includeCompletedNew"] ?? "true") == "true")
                categories.Add('-');
            if ((config["includeCompletedOld"] ?? "true") == "true")
                categories.Add('.');

            // Calculate the age of each not-completed and completed message.
            var now = DateTime.Now;
            var histogram = new Histogram(categories);
            var countFlagged = 0;
            var countNewFlagged = 0;
            var countNewComplete = 0;

            FindFlaggedMessages(service, message =>
            {
                try
                {
                    if (message.Flag.DueDate.Year > 1 || message.Flag.CompleteDate.Year > 1)
                    {
                        var messageAge = (now - message.DateTimeReceived).TotalDays / 7;
                        if (message.Flag.FlagStatus == ItemFlagStatus.Flagged)
                        {
                            if ((now - message.DateTimeReceived).TotalDays >= 7)
                                histogram.Add('#', messageAge);
                            else
                                histogram.Add('+', messageAge);
                        }
                        else if (message.Flag.FlagStatus == ItemFlagStatus.Complete)
                        {
                            if ((now - message.Flag.CompleteDate).TotalDays < 7)
                                histogram.Add('-', messageAge);
                            else
                                histogram.Add('.', messageAge);
                        }
                    }

                    if (message.Flag.FlagStatus == ItemFlagStatus.Flagged || message.Flag.FlagStatus == ItemFlagStatus.Complete)
                    {
                        if ((now - message.DateTimeReceived).TotalDays < 7)
                            countNewFlagged++;
                    }
                    if (message.Flag.FlagStatus == ItemFlagStatus.Flagged)
                    {
                        countFlagged++;
                    }
                    else if (message.Flag.FlagStatus == ItemFlagStatus.Complete)
                    {
                        if ((now - message.Flag.CompleteDate).TotalDays < 7)
                            countNewComplete++;
                    }
                }
                catch (ServiceObjectPropertyException)
                {
                }
            });

            Console.WriteLine($"Flagged:  {countFlagged,3} ( +{countNewFlagged} -{countNewComplete} => {countNewFlagged - countNewComplete:+#;-#;0} )");

            var countCategories = new List<char>(4);
            if ((config["countFlaggedOld"] ?? "true") == "true")
                countCategories.Add('#');
            if ((config["countFlaggedNew"] ?? "true") == "true")
                countCategories.Add('+');
            if ((config["countCompletedNew"] ?? "false") == "true")
                countCategories.Add('-');
            if ((config["countCompletedOld"] ?? "false") == "true")
                countCategories.Add('.');

            var output = new HistogramOutput()
            {
                BinSize = int.Parse(config["binSize"] ?? "0"),
                Width = int.Parse(config["width"] ?? "0") - 16,
                Height = int.Parse(config["height"] ?? "0"),
            };

            histogram.RenderTo(output, countCategories);

            Console.WriteLine("Weeks   | Flg | Flagged #/+  Complete -/.");
            for (var i = 0; i < output.Graph.Length; i++)
            {
                // Everything before {3} comes to 16 characters, the adjustment used above.
                Console.WriteLine("{0,3}-{1,3} | {2,3} | {3}",
                    output.Base + output.BinSize * i,
                    output.Base + output.BinSize * (i + 1) - 1,
                    output.Values[i],
                    output.Graph[i]
                );
            }
        }

        private static void FindFlaggedMessages(ExchangeService service, Action<Item> onMessage)
        {
            var PidTagFolderType = new ExtendedPropertyDefinition(0x3601, MapiPropertyType.Integer);
            var PidTagFlagStatus = new ExtendedPropertyDefinition(0x1090, MapiPropertyType.Integer);

            // Find Outlook's own search folder "AllItems", which includes all folders in the account.
            var allItemsView = new FolderView(10);
            var allItems = service.FindFolders(WellKnownFolderName.Root,
                new SearchFilter.SearchFilterCollection(LogicalOperator.And) {
                    new SearchFilter.IsEqualTo(PidTagFolderType, "2"),
                    new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems"),
                }, allItemsView);

            if (allItems.Folders.Count != 1)
            {
                throw new MissingMemberException("AllItems");
            }

            // Find all items that are flagged.
            var flaggedFilter = new SearchFilter.Exists(PidTagFlagStatus);
            var flaggedView = new ItemView(1000)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.DateTimeReceived, ItemSchema.Flag),
                Traversal = ItemTraversal.Shallow,
            };

            FindItemsResults<Item> flagged;
            do
            {
                flagged = allItems.Folders[0].FindItems(flaggedFilter, flaggedView);
                foreach (var item in flagged.Items)
                {
                    onMessage(item);
                }
                flaggedView.Offset = flagged.NextPageOffset ?? 0;
            } while (flagged.MoreAvailable);
        }
    }
}
