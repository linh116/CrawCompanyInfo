package org.wesol.helper;

import org.apache.commons.cli.*;

public class ClientOptionHelper {

    public static CommandLine createCommandLineOption(String[] args){
        Options options = new Options();

        Option input = new Option("d", "driver", true, "driver file path");
        input.setRequired(true);
        options.addOption(input);

        Option output = new Option("o", "output", true, "output folder");
        output.setRequired(true);
        options.addOption(output);

        Option excelFile = new Option("e", "excel", true, "excel file name");
        excelFile.setRequired(true);
        options.addOption(excelFile);

        Option stopPage = new Option("p", "stoppage", true, "Stop after X page crawled nothing. Default is unlimited.");
        stopPage.setRequired(false);
        options.addOption(stopPage);

        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        try {
            CommandLine cmd = parser.parse(options, args);
            return cmd;
        } catch (ParseException e) {
            System.out.println(e.getMessage());
            System.out.println("Your args: ");
            for (String arg : args){
                System.out.print(arg + "\t");
            }
            System.out.println();
            formatter.printHelp("CrawCompanyInfo", options);

            System.exit(1);
        }
        return null;
    }
}
