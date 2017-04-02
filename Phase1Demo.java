package com.him;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class Phase1Demo {

	private static Map<Integer,String[]> defectDetailsMap=new TreeMap<Integer,String[]>();
	private static Map<Integer,String[]> mailDetailsMap=new TreeMap<Integer,String[]>();
	private static Map<Integer,String[]> highlightsDetailsMap=new TreeMap<Integer,String[]>();
	private static Map<Integer,String[]> piDetailsMap=new TreeMap<Integer,String[]>();
	private static Map<Integer,String[]> monthlyTicketDetailsMap=new TreeMap<Integer,String[]>();
	private static Map<Integer,String[]> categoryViseTicketDetailsMap=new TreeMap<Integer,String[]>();
	private static Map<Integer,String[]> competencyDetailsMap=new TreeMap<Integer,String[]>();
	private static Map<Integer,String[]> barGraphDetailsMap=new TreeMap<Integer,String[]>();
	
	/**
	 * @param args
	 * @throws Exception 
	 */
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		XWPFDocument doc= new XWPFDocument();
		FileOutputStream out = new FileOutputStream(new File("C:\\L3ProjectFile\\Dashboard.docx"));
		int totalSheet=readFileData();
		writeFileData(doc,totalSheet);
		doc.write(out);
		out.close();
		System.out.println("Dashboard.docx written successfully");
	}
	public static int readFileData() throws Exception
	{
		System.out.println("File Reading Started"); 
        //Create Workbook instance holding reference to .xlsx file
		int totalNumberOfSheet=0,valueOfSheet=0;
     	try 
     	{
     			String str=null;
				int k=1;int keyValue=1;int srNumber=1;
				FileInputStream file = new FileInputStream(new File("C:\\L3ProjectFile\\L3CompleteExcel.xlsx"));
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				totalNumberOfSheet=workbook.getNumberOfSheets();
				System.out.println("Total Sheet:: "+totalNumberOfSheet);
				for(int i=0;i<totalNumberOfSheet;i++)
				{
					System.out.println("************ Sheet "+(i+1)+"************** start");
					XSSFSheet sheet = workbook.getSheetAt(i);
					//Iterate through each rows one by one
					Iterator<Row> rowIterator = sheet.iterator();
					while (rowIterator.hasNext()) 
					{
						 Row row = rowIterator.next();
						 //For each row, iterate through all the columns
						 if(row.getRowNum()==0){
							   continue; //just skip the rows if row number is 0 or 1
						 }
						Iterator<Cell> cellIterator = row.cellIterator();
						String obj[]=new String[100];
						obj[0]=""+srNumber;
						while (cellIterator.hasNext()) 
						{
								Cell cell = cellIterator.next();
								//Check the cell type and format accordingly
								switch (cell.getCellType()) 
								{
									case Cell.CELL_TYPE_NUMERIC:
										obj[k]=String.valueOf(cell.getNumericCellValue());
										k++;
										System.out.print(cell.getNumericCellValue() + "\t");
										break;
									case Cell.CELL_TYPE_STRING:
										str=cell.getStringCellValue();
										obj[k]=str;
										k++;
										System.out.print(str + "\t");
										break;
								}
						}//Cell End
						if(valueOfSheet==0){
								if(obj[0]!=null && obj[1]!=null && obj[2]!=null && obj[3]!=null && obj[4]!=null){
								defectDetailsMap.put(keyValue,obj);
								}
						}
						else if(valueOfSheet==1){
							if(obj[0]!=null && obj[1]!=null && obj[2]!=null){
								mailDetailsMap.put(keyValue,obj);
								}
							}
						else if(valueOfSheet==2){
							if(obj[0]!=null && obj[1]!=null && obj[2]!=null){
								highlightsDetailsMap.put(keyValue,obj);
							}
						}
						else if(valueOfSheet==3){
							if(obj[0]!=null && obj[1]!=null && obj[2]!=null){
									monthlyTicketDetailsMap.put(keyValue,obj);
							}
						}
						else if(valueOfSheet==4){
							if(obj[0]!=null && obj[1]!=null && obj[2]!=null){
									piDetailsMap.put(keyValue,obj);
							}
						}
						k=1;
						keyValue++;
						srNumber++;
						System.out.println("");
					}//Row End
					System.out.println("************ Sheet "+(i+1)+"************** End\n");
					valueOfSheet++;
					keyValue=1;
					srNumber=1;
				}//Sheet End
				file.close();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return totalNumberOfSheet;
	}
	
	
	public static void writeFileData(XWPFDocument doc,int totalsheet)
	{
		for(int i=0;i<totalsheet;i++)
		{
			createHeader(doc,i);
			createStyledTable(doc,i);
		}
	}
	
	public static void createHeader(XWPFDocument doc,int value){
		XWPFParagraph paragraph = doc.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run=paragraph.createRun();
		run.setFontSize(15);
		run.setColor("0000A0");
		run.setBold(true);
		  
		if(value==0){
			run.setText("L3 Defect Details\n");
		}
		else if(value==1){
			run.setText("L3 Mail Issues\n");
		}
		else if(value==2){
			run.setText("L3 Highlights\n");
		}
		else if(value==3){
			run.setText("L3 Monthly Ticket Count Details\n");
		}
		else if(value==4){
			run.setText("L3-PI Details\n");
		}
		
	}
	public static void createStyledTable(XWPFDocument doc,int value){
		List convertedList=null;
		int size=0;
		if(value==0){
			size=defectDetailsMap.size();
			convertedList=convertMapDataToList(defectDetailsMap,value);
		}
		else if(value==1){
			size=mailDetailsMap.size();
			convertedList=convertMapDataToList(mailDetailsMap,value);
		}
		else if(value==2){
			size=highlightsDetailsMap.size();
			convertedList=convertMapDataToList(highlightsDetailsMap,value);
		}
		else if(value==3){
			size=monthlyTicketDetailsMap.size();
			convertedList=convertMapDataToList(monthlyTicketDetailsMap,value);
		}
		else if(value==4){
			size=piDetailsMap.size();
			convertedList=convertMapDataToList(piDetailsMap,value);
		}
		
		
		// Create a new table with 6 rows and 3 columns
    	int nRows = size+1;
    	int nCols =((String[])convertedList.get(0)).length;
    	System.out.println("Total Row :: "+nRows+" cell:: "+nCols);
    	
    	XWPFTable table = doc.createTable(nRows, nCols);

        // Set the table style. If the style is not defined, the table style
        // will become "Normal".
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CTString styleStr = tblPr.addNewTblStyle();
        styleStr.setVal("StyledTable");

        // Get a list of the rows in the table
        List<XWPFTableRow> rows = table.getRows();
        int rowCt = 0;
        int colCt = 0;
        int arr=0;
              
        for (XWPFTableRow row : rows) {
        	// get table row properties (trPr)
        	CTTrPr trPr = row.getCtRow().addNewTrPr();
        	// set row height; units = twentieth of a point, 360 = 0.25"
        	CTHeight ht = trPr.addNewTrHeight();
        	ht.setVal(BigInteger.valueOf(360));
        	
        	
        	String[] objArray=(String[]) convertedList.get(rowCt);
	        
        	
        	// get the cells in this row
        	List<XWPFTableCell> cells = row.getTableCells();
            // add content to each cell
        		for (XWPFTableCell cell : cells) {
        		// get a table cell properties element (tcPr)
        		CTTcPr tcpr = cell.getCTTc().addNewTcPr();
        		// set vertical alignment to "center"
        		CTVerticalJc va = tcpr.addNewVAlign();
        		va.setVal(STVerticalJc.CENTER);

        		// create cell color element
        		CTShd ctshd = tcpr.addNewShd();
                ctshd.setColor("auto");
                ctshd.setVal(STShd.CLEAR);
                if (rowCt == 0) {
                	// header row
                	ctshd.setFill("A7BFDE");
                }
            	else if (rowCt % 2 == 0) {
            		// even row
                	ctshd.setFill("D3DFEE");
            	}
            	else {
            		// odd row
                	ctshd.setFill("EDF2F8");
            	}

                // get 1st paragraph in cell's paragraph list
                XWPFParagraph para = cell.getParagraphs().get(0);
                // create a run to contain the content
                XWPFRun rh = para.createRun();
                if (rowCt == 0) {
                	// header row
                	rh.setFontSize(13);
                    rh.setText(objArray[arr]);
                	rh.setBold(true);
                    para.setAlignment(ParagraphAlignment.CENTER);
                }
            	else {
            		rh.setText(objArray[arr]);
                    para.setAlignment(ParagraphAlignment.LEFT);
            	}
                colCt++;
                arr++;
        	} // for cell
        	colCt = 0;
        	arr=0;
        	rowCt++;
       	}//for row
		
        XWPFParagraph paragraph = doc.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run=paragraph.createRun();
		run.setText("\n\n\n\n");		
	}
	
	
	
	public static List<String []> convertMapDataToList(Map<Integer,String[]> data,int value)
	{
		List<String[]> totalList=new ArrayList<String[]>();
		String str[]=null;
		Map.Entry entry;
		if(value==0){
			str=new String[]{"S/No","Defect_Id","Issue Description","Solved By","Ticket count"};
		}
		else if(value==1){
			str=new String[]{"S/No","Mail Issue","Solved By"};
		}
		else if(value==2){
			str=new String[]{"S/No","Highlights","Solved By"};
		}
		else if(value==3){
			str=new String[]{"S/No","Month","Ticket Count"};
		}
		else if(value==4){
			str=new String[]{"S/No","PI","Solved By"};
		}
		
		totalList.add(str);
		for(Object obj : data.entrySet())
		{
			entry = (Map.Entry)obj;
			totalList.add((String[]) entry.getValue());
		}
		for(int i=1;i<totalList.size();i++)
		{
			String str1[]=totalList.get(i);
			for(int j=0;j<str1.length;j++)
			{
				System.out.print(" value:: "+str1[j]);
			}
			System.out.println();
		}
		return totalList;
	}
}
