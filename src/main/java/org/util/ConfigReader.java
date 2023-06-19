package org.util;

import org.testng.models.Configs;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ConfigReader {

    private static String configProName = "config.properties";
    private static Properties properties = new Properties();
    public static Configs configs = new Configs();
    private static Properties readProperties(){
        if(!properties.isEmpty())
            return properties;

        try {
            InputStream inputStream = ConfigReader.class.getResourceAsStream("/" + configProName);
            properties.load(inputStream);
            //properties.forEach((key, value) -> System.out.println("Key : " + key + ", Value : " + value));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return properties;
    }

    public static Configs getAndSetConfigs(){
        configs.setMainXLSXFilePath(readProperties().getProperty("xlsx.file.location"));
        configs.setSheet1Name(readProperties().getProperty("xlsx.sheet1.name"));
        configs.setSheet2Name(readProperties().getProperty("xlsx.sheet2.name"));

        configs.setPrintOutput(Boolean.valueOf(readProperties().getProperty("result.print.output.to.console")).booleanValue());

        configs.setPrimaryKeyColHeader(readProperties().getProperty("xlsx.primary.key.column.header"));

        configs.setSheet1ResultColHeader(readProperties().getProperty("xlsx.sheet1.result.column.header"));
        configs.setSheet1ResultCmntColHeader(readProperties().getProperty("xlsx.sheet1.result.comments.column.header"));

        configs.setSheet1ResultColNum(validateInt("xlsx.sheet1.result.column.number",
                readProperties().getProperty("xlsx.sheet1.result.column.number")) - 1);
        configs.setSheet1ResultCmntColNum(validateInt("xlsx.sheet1.result.comments.column.number",
                readProperties().getProperty("xlsx.sheet1.result.comments.column.number")));

        configs.setSheet2ValueColHeader(readProperties().getProperty("xlsx.sheet2.value.column.header"));


        return configs;
    }

    public static int validateInt(String name, String value){
        int x = -9999;
        try{
            x = Integer.parseInt(value);
            if(x < 1)
                throw new NumberFormatException();
        }catch (NumberFormatException ne){
            System.out.println("\n\nError: " +name + " in " + configProName + " should be in number format\n");
            ne.printStackTrace();
            System.exit(0);
        }
        return x;
    }

}
