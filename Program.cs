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

            var debug = new CLP.Arguments.SwitchArgument('d', "debug", false);

            var commandLineParser = new CLP.CommandLineParser()
            {
                Arguments = {
                    config,
                    debug,
                }
            };

            try
            {
                commandLineParser.ParseCommandLine(args);

                InnerMain(new ConfigurationBuilder()
                    .AddJsonFile(config.Value.FullName, true)
                    .Build(), debug.Value);
            }
            catch (CLP.Exceptions.CommandLineException e)
            {
                Console.WriteLine(e.Message);
            }
        }

        static void InnerMain(IConfigurationRoot config, bool debug)
        {
            var service = new ExchangeService(ExchangeVersion.Exchange2013);

            LogIn(config.GetSection("credentials"), service);

            Histogram(config.GetSection("histogram"), service, debug);
        }

        private static void LogIn(IConfigurationSection config, ExchangeService service)
        {
            service.Credentials = new WebCredentials(config["username"], config["password"]);
            service.AutodiscoverUrl(config["email"], redirectionUri =>
                new Uri(redirectionUri).Scheme == "https"
            );
        }

        private static void Histogram(IConfigurationSection config, ExchangeService service, bool debug)
        {
            var categories = new List<char>(4);
            if ((config["includeFlaggedOld"] ?? "True") == "True")
                categories.Add('#');
            if ((config["includeFlaggedNew"] ?? "True") == "True")
                categories.Add('+');
            if ((config["includeCompletedNew"] ?? "True") == "True")
                categories.Add('-');
            if ((config["includeCompletedOld"] ?? "True") == "True")
                categories.Add('.');

            // Calculate the age of each not-completed and completed message.
            var now = DateTimeOffset.Now;
            now = now.AddDays(1 - now.TimeOfDay.TotalDays);
            var histogram = new Histogram(categories);
            var separateFlaggedCompleted = (config["separateFlaggedCompleted"] ?? "False") == "True";
            var daysPerBin = uint.Parse(config["daysPerBin"] ?? "7");
            var countFlagged = 0;
            var countNewFlagged = 0;
            var countNewComplete = 0;

            if (debug)
                Console.WriteLine("Now:      {0}", now.DateTime);

            FindFlaggedMessages(service, message =>
            {
                try
                {
                    var messageAge = (now - message.DateTimeReceived).TotalDays;
                    var completedAge = (now - message.Flag.CompleteDate).TotalDays - 0.5;

                    if (debug && messageAge < 7)
                    {
                        Console.WriteLine(message.Flag.FlagStatus == ItemFlagStatus.Flagged ?
                            "Received: {0} ({1,3:F1} days ago)  Flag: {2,-8}" :
                            "Received: {0} ({1,3:F1} days ago)  Flag: {2,-8}  Completed: {3} ({4,3:F1} days ago)",
                            message.DateTimeReceived,
                            messageAge,
                            message.Flag.FlagStatus,
                            message.Flag.CompleteDate,
                            completedAge
                        );
                    }

                    if (message.Flag.FlagStatus == ItemFlagStatus.Flagged)
                    {
                        if (messageAge >= 7)
                            histogram.Add('#', messageAge / daysPerBin);
                        else
                            histogram.Add('+', messageAge / daysPerBin);
                    }
                    else if (message.Flag.FlagStatus == ItemFlagStatus.Complete)
                    {
                        if (separateFlaggedCompleted)
                        {
                            if (messageAge >= 7)
                                histogram.Add('#', messageAge / daysPerBin);
                            else
                                histogram.Add('+', messageAge / daysPerBin);
                        }
                        if (completedAge < 7)
                            histogram.Add('-', (separateFlaggedCompleted ? completedAge : messageAge) / daysPerBin);
                        else
                            histogram.Add('.', (separateFlaggedCompleted ? completedAge : messageAge) / daysPerBin);
                    }

                    if (message.Flag.FlagStatus == ItemFlagStatus.Flagged || message.Flag.FlagStatus == ItemFlagStatus.Complete)
                    {
                        if (messageAge < 7)
                            countNewFlagged++;
                    }
                    if (message.Flag.FlagStatus == ItemFlagStatus.Flagged)
                    {
                        countFlagged++;
                    }
                    else if (message.Flag.FlagStatus == ItemFlagStatus.Complete)
                    {
                        if (completedAge < 7)
                            countNewComplete++;
                    }
                }
                catch (ServiceObjectPropertyException error)
                {
                    Console.WriteLine("Error: {0}: {1}", error.PropertyDefinition, error.Message);
                }
            });

            Console.WriteLine($"Flagged:  {countFlagged,3} ( +{countNewFlagged} -{countNewComplete} => {countNewFlagged - countNewComplete:+#;-#;0} )");

            if (debug)
                return;

            var countCategories = new List<char>(4);
            if ((config["countFlaggedOld"] ?? "True") == "True")
                countCategories.Add('#');
            if ((config["countFlaggedNew"] ?? "True") == "True")
                countCategories.Add('+');
            if ((config["countCompletedNew"] ?? "False") == "True")
                countCategories.Add('-');
            if ((config["countCompletedOld"] ?? "False") == "True")
                countCategories.Add('.');

            var countNegCategories = new List<char>(4);
            if ((config["countNegFlaggedOld"] ?? "False") == "True")
                countNegCategories.Add('#');
            if ((config["countNegFlaggedNew"] ?? "False") == "True")
                countNegCategories.Add('+');
            if ((config["countNegCompletedNew"] ?? "False") == "True")
                countNegCategories.Add('-');
            if ((config["countNegCompletedOld"] ?? "False") == "True")
                countNegCategories.Add('.');

            var output = new HistogramOutput()
            {
                BinSize = int.Parse(config["binSize"] ?? "0"),
                Width = int.Parse(config["width"] ?? "0") - 17,
                Height = int.Parse(config["height"] ?? "0"),
            };
            if (config["minScale"] != null)
                output.MinScale = double.Parse(config["minScale"]);
            if (config["maxScale"] != null)
                output.MaxScale = double.Parse(config["maxScale"]);

            histogram.RenderTo(output, countCategories, countNegCategories);

            if (daysPerBin == 1)
                Console.WriteLine("Days    |  Num | Flagged #/+  Complete -/.");
            else if (daysPerBin == 7)
                Console.WriteLine("Weeks   |  Num | Flagged #/+  Complete -/.");
            else
                Console.WriteLine("{0,2} days |  Num | Flagged #/+  Complete -/.", daysPerBin);
            for (var i = 0; i < output.Graph.Length; i++)
            {
                // Everything before {3} comes to 17 characters, the adjustment used above.
                Console.WriteLine("{0,3}-{1,3} | {2,4} | {3}",
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
                new SearchFilter.SearchFilterCollection(LogicalOperator.And)
                {
                    new SearchFilter.IsEqualTo(PidTagFolderType, "2"),
                    new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems"),
                },
                allItemsView);

            if (allItems.Folders.Count != 1)
            {
                throw new MissingMemberException("AllItems");
            }

            // Find the Funk folder.
            var junkFolder = Folder.Bind(service, WellKnownFolderName.JunkEmail);

            // Find all items that are flagged and not in the Junk folder.
            var flaggedFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And)
            {
                new SearchFilter.Exists(PidTagFlagStatus),
                new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"),
                new SearchFilter.IsNotEqualTo(ItemSchema.ParentFolderId, junkFolder.Id.UniqueId),
            };
            var flaggedView = new ItemView(1000)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.DateTimeReceived, ItemSchema.Flag),
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
