package main;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.internal.FileHelper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CountWorkTime {

	static int ot = 0;
	static List<Integer> ots = new ArrayList<Integer>();
	private static DecimalFormat df2 = new DecimalFormat("#.##");
	float hours = 0.00f;
	float days = 0.00f;
	static int sums = 0;

	public static void main(String[] args) throws ParseException {

		final List<String> result = getFilePathToSaveStatic();

		for (String part : result) {
			try {
				countDqyWork(part);
			} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
				e.printStackTrace();
			}
		}
		for (int sum : ots) {
			sums += sum;
		}

		if (((sums / (double) 60) / (double) 8) > 1) {
			//1.94 = 1 วัน 1 ชั่วโมง 34 นาที
//			double xx = ((sums / (double) 60) / (double) 8);
			System.out.println(sums / 8);
		} else {
			
		}
		System.out.println("รวมหมดทุกเดือนได้วันหยุดเพิ่ม : " + df2.format( (sums / (double) 60) / (double) 8 ) + " วัน:ชั่วโมง");

	}

	public static void countDqyWork(String part)
			throws IOException, EncryptedDocumentException, InvalidFormatException, ParseException {
		XSSFWorkbook workbook = new XSSFWorkbook(new File(part));
		XSSFSheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					if (cell.getNumericCellValue() != 0) {
						double x = cell.getNumericCellValue();
						CellAddress cellAddress = cell.getAddress();
						String strCellAddress = cellAddress.toString();
						String xx = Double.toString(x);
						String findC = xx + ":" + strCellAddress;
						if (findC.indexOf("C") > 0) {
							Date date = cell.getDateCellValue();
							SimpleDateFormat sdf = new SimpleDateFormat("mm");
							ot += Integer.parseInt(sdf.format(date));
						}
					}
					break;
				}
			}
		}
		String[] temp = part.split("/");
		String month = "";
		for (int i = 0; i < temp.length; i++) {
			if (i == 6) {
				String subStr = temp[i].indexOf(".") > 0 ? temp[i].toString().substring(0, 7) : temp[i];
				String[] ym = subStr.split("_");
				SimpleDateFormat d = new SimpleDateFormat("yyyy - MMM");
				Calendar calendar = new GregorianCalendar(Integer.parseInt(ym[0]), Integer.parseInt(ym[1]) -1, 1);
				month = d.format(calendar.getTime());
			}

		}
		System.out.println("===== " + month + " =====");
		System.out.println("ทำงานล่วงเวลาทั้งหมด : " + ot + " นาที");
		System.out.println("แลกเป็นวันหยุดได้ : " + ot / (double) 60 + " ชั่วโมง");
		System.out.println("==========================");
		ots.add(ot);
		ot = 0;
	}

	public static List<String> getFilePathToSaveStatic() {
		Properties prop = new Properties();
		List<String> result = new ArrayList<String>();
		try (InputStream inputStream = FileHelper.class.getClassLoader().getResourceAsStream("config.properties")) {
			prop.load(inputStream);
			result.add(prop.getProperty("json.filepath1"));
			result.add(prop.getProperty("json.filepath2"));
			result.add(prop.getProperty("json.filepath3"));
		} catch (IOException e) {
			e.printStackTrace();
		}
		return result;

	}
}
