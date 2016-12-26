package scrapping;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class TestScrap {

	public static void main(String[] args) throws IOException {
		// Jsoup parser
		URL url = new URL("https://en.wikipedia.org/wiki/List_of_people_who_died_climbing_Mount_Everest");
		Document doc = Jsoup.parse(url, 3000);

		// Apache POI excel Document creation
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet("Mount Everest Death Information ");

		// Create a first row and set value for headings
		XSSFRow row1 = spreadsheet.createRow(0);
		XSSFCell r1c1 = row1.createCell(0);
		r1c1.setCellValue("Name");

		XSSFCell r1c2 = row1.createCell(1);
		r1c2.setCellValue("Death Year");

		XSSFCell r1c3 = row1.createCell(2);
		r1c3.setCellValue("Country");

		// select the table and rows count
		Element table = doc.select("tbody").get(0);
		Elements rows = table.select("tr");

		// looping in rows to get information

		for (int i = 1; i < rows.size(); i++) {
			Element row = rows.get(i);
			Elements cols = row.select("td");

			// Parsing country as per class string value
			String country = cols.select("a[title]").text();

			// Parsing name and death date as per td value
			String name = cols.get(0).text();
			String date = cols.get(1).text();

			// Substring for date
			String a = date.substring(0, 10);

			// Writing in Excel file
			XSSFRow row2 = spreadsheet.createRow(i);
			XSSFCell rxc1 = row2.createCell(0);
			XSSFCell rxc2 = row2.createCell(1);
			XSSFCell rxc3 = row2.createCell(2);
			rxc1.setCellValue(name);
			rxc2.setCellValue(a);
			rxc3.setCellValue(country);
		}

		FileOutputStream out = new FileOutputStream("Mt. Everest Deaths.xlsx");
		workbook.write(out);
		out.close();
		System.out.println("typesofcells.xlsx written successfully");

	}

}
