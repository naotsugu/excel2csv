package com.mammb.excel2csv;

import java.io.File;
import java.nio.file.Paths;

public class Main {

    public static void main(String[] args) {

        if (args.length < 1) {
            System.err.println("please provide an excel file to convert.");
            System.exit(1);
        }

        final File xlsx = new File(args[0]);
        if (!xlsx.exists() || xlsx.isDirectory() || !xlsx.getPath().endsWith(".xlsx")) {
            System.err.println("not found a file: " + xlsx.getPath());
            System.exit(1);
        }

        try {
            Excel2Csv excel2Csv = new Excel2Csv(
                    xlsx.toPath(),
                    Paths.get(xlsx.getPath().replace(".xlsx", ".csv")),
                    (args.length > 1) ? args[1] : "");
            excel2Csv.process();
        } catch (Exception e) {
            System.err.println("error. " + e);
            System.exit(1);
        }
        System.exit(0);
    }

}
