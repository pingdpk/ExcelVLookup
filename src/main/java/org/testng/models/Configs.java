package org.testng.models;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class Configs {
    private String mainXLSXFilePath;
    private boolean printOutput;
    private String sheet1Name;
    private String sheet2Name;

    private String primaryKeyColHeader;
    private int sheet1ResultColNum;
    private String sheet1ResultColHeader;
    private int sheet1ResultCmntColNum;
    private String sheet1ResultCmntColHeader;

    private String sheet2ValueColHeader;

}
