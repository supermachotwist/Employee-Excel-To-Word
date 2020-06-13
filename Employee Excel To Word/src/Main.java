import java.io.*;
import java.util.*;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;

public class Main {

    public static void main(String[] args) throws Exception{
    	//Picture folder absolute path
        Scanner stdin = new Scanner(System.in);
        System.out.print("Enter Absolute path to the picture folder:");
        String picturePath = stdin.nextLine();
        //excel file absolute path
        System.out.print("Enter the name of the excel folder (no extension):");
    	
        XWPFDocument doc = new XWPFDocument();
        FileInputStream file = new FileInputStream(new File(stdin.nextLine() + ".xlsx"));
        FileOutputStream out = new FileOutputStream(new File("Employee Output.docx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        
        stdin.close();
        
        //picture directory
        File directory = new File(picturePath);
        String[] fileNames = directory.list();
        List<String> filesNameList = Arrays.asList(fileNames);
        //File[] files = directory.listFiles();
        //List<File> filesList = Arrays.asList(files);
        
        //initialize paragraph and run
        XWPFParagraph paragraph;
        XWPFRun run;
        XWPFParagraph docParagraph;
        XWPFRun docRun;
        
        //createTable(row, column)
        XWPFTable table = null;
        //current column 
        int column = 0;
        //current row
        int docRow = 0;
        //current department
        String department = "";
        DataFormatter dataFormatter = new DataFormatter();
        
        System.out.println("Running. This may take a minute.");
        
        for (Row row : sheet) {
        	if (row.getRowNum() == 0) {
        		continue;
        	}
        	//wrap around table
        	if (column >= 5) {
        		column = 0;
        		docRow++;
        		if (docRow >= 3) {
        			docRow = 0;
        		}
        	}
        	
        	//add title of department and create new table
        	if ((docRow == 0 && column == 0) || (department != row.getCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue())) {
        		department = row.getCell(3).getStringCellValue();
        		docParagraph = doc.createParagraph();
        		docRun = docParagraph.createRun();
        		docParagraph.setAlignment(ParagraphAlignment.CENTER);
        		docRun.setFontSize(36);
        		docRun.setText(row.getCell(3).getStringCellValue());
        		table = doc.createTable(3,5);
        		table.removeBorders();
        		table.setTableAlignment(TableRowAlign.CENTER);;
        		docParagraph = doc.createParagraph();
        		docRun = docParagraph.createRun();
        		docRun.addBreak(BreakType.PAGE);
        		column = 0;
        		docRow = 0;
        	}
        	
        	//set width of cell
        	table.getRow(docRow).getCell(column).setWidth("2000");
        	//Fill in table
        	paragraph = table.getRow(docRow).getCell(column).getParagraphs().get(0);
        	paragraph.setAlignment(ParagraphAlignment.CENTER);
        	run = paragraph.createRun();
        	//insert picture and ID
        	if (row.getCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK).getCellType().equals(CellType.STRING)) {
        		//filter files starting with employee ID
        		File[] files = directory.listFiles(new FilenameFilter() {
        			public boolean accept(File directory, String name) {
        				return name.startsWith(row.getCell(1).getStringCellValue());
        			}
        		});
        		if (files.length != 0 && filesNameList.contains(files[0].getName())) {
        			run.addPicture(new FileInputStream(new File(picturePath + "\\" +  files[0].getName())), Document.PICTURE_TYPE_JPEG, files[0].getName(), Units.toEMU(60), Units.toEMU(72));
        			run.addBreak();
        		}
        		//default picture if one doesn't exist
        		else {
        			run.addPicture(new FileInputStream(new File("DefaultPicture.png")), Document.PICTURE_TYPE_PNG, "DefaultPicture.png", Units.toEMU(60), Units.toEMU(72));
        			run.addBreak();
        		}
        		run.setText(row.getCell(1).getStringCellValue());
        		run.addBreak();
        	}
        	//insert name
        	if (row.getCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK).getCellType().equals(CellType.STRING)) {
        		run.setText(row.getCell(0).getStringCellValue());
        		run.addBreak();
        	}
        	//insert job/Section
        	if (row.getCell(13, MissingCellPolicy.CREATE_NULL_AS_BLANK).getCellType().equals(CellType.NUMERIC)) {
        		run.setText(dataFormatter.formatCellValue(row.getCell(13)));
        		run.addBreak();
        	}
        	//insert date of birth
        	if (row.getCell(9, MissingCellPolicy.CREATE_NULL_AS_BLANK).getCellType().equals(CellType.NUMERIC)) {
        		run.setText(dataFormatter.formatCellValue(row.getCell(9)));
        		run.addBreak();
        	}
        	//insert date of employment
        	if (row.getCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK).getCellType().equals(CellType.STRING)) {
        		run.setText(dataFormatter.formatCellValue(row.getCell(4)));
        		run.addBreak();
        	}
        	column++;
        }
        file.close();
        doc.write(out);
        workbook.close();
        doc.close();
        System.out.print("Program Complete. Check current folder for output file.");
    }
}