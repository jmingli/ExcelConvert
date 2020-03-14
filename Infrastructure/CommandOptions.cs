using System;
using System.Collections.Generic;
using ExcelConvert.Extensions;
using Microsoft.Extensions.CommandLineUtils;

namespace ExcelConvert.Infrastructure
{
    public class CommandOptions
    {
        public enum ConvertOption
        {
            QuickBookBalanceSheet,
        }

        //option
        public CommandOption Convert { get; set; } = null!;

        //arguments
        public CommandOption InFile { get; set; } = null!;
        public CommandOption OutFile { get; set; } = null!;


        public static CommandOptions ParseAndCreate(CommandLineApplication cmdApp)
        {
            var commandOptions = new CommandOptions
            {
                //options
                Convert = cmdApp.Option(
                    "--convert <name>", $"Convert. {GetEnumOptions<ConvertOption>()}",
                    CommandOptionType.SingleValue),

                InFile = cmdApp.Option(
                    "--i", "In File Name\n",
                    CommandOptionType.SingleValue),

                OutFile = cmdApp.Option(
                    "--o", "Out File Name\n",
                    CommandOptionType.SingleValue),
            };

            return commandOptions;
        }

        private static string GetEnumOptions<T>()
        {
            var list = new List<string>();

            foreach (var item in Enum.GetValues(typeof(T)))
            {
                list.Add(item!.ToString()!.ToSnakeCase());
            }

            var delimiter = $"{Environment.NewLine}        ";
            var options = string.Join(delimiter, list);

            return $"{Environment.NewLine}{delimiter}{options}{Environment.NewLine}";
        }
    }
}
