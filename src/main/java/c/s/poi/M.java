package c.s.poi;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class M {

	public static void main(String[] args) throws IOException {

		byte[] original = Files.readAllBytes(Paths.get("/tmp/x.xlsx"));

		loop(original);

		long s = 0;

		int loops = 0;

		for (int i = 0; i < loops; i++) {

			s += loop(original);
		}

		if(loops!=0)
		System.out.println(new BigDecimal(s).divide(new BigDecimal(loops), RoundingMode.HALF_UP));

	}

	static long loop(byte[] original) throws IOException {

		long l = System.currentTimeMillis();

		byte[] buf = original.clone();

		XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(buf));

		

		final AtomicInteger maxrow=new AtomicInteger(0);
		
		final AtomicInteger maxcol=new AtomicInteger(0);
		

		wb.sheetIterator().forEachRemaining(s -> {

			
			s.rowIterator().forEachRemaining(r ->{
				
				
				
				if(maxrow.get()<r.getRowNum()) {
					maxrow.set(r.getRowNum());
				}
				
				
				if(maxcol.get()<r.getLastCellNum()) {
					maxcol.set(r.getLastCellNum());
				}
				
				//r.getLastCellNum()
				
			});
			
			
			// System.out.println();
		});
		

		wb.sheetIterator().forEachRemaining(s -> {
			
			
			XSSFName name = wb.createName();

			name.setNameName(s.getSheetName());

			System.out.println(s.getSheetName() + "!A1:" + CellReference.convertNumToColString(maxcol.get()-1)+  ""+(maxrow.get()+1));
			
			name.setRefersToFormula(s.getSheetName() + "!A1:" + CellReference.convertNumToColString(maxcol.get()-1)+  ""+(maxrow.get()+1));

		});
		/*
		 * 
		 * 
		 * 
		 */

		CellReference cra = new CellReference("A1");

		CellReference crb = new CellReference("B1");

		CellReference crc1 = new CellReference("C1");

		CellReference crc = new CellReference("A1");

		if (wb.getSheetAt(0).getRow(cra.getRow()) == null)

		{
			wb.getSheetAt(0).createRow(cra.getRow());
		}

		wb.getSheetAt(0).getRow(cra.getRow()).

				getCell(cra.getCol(), MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("1");

		wb.getSheetAt(0).getRow(crb.getRow()).getCell(crb.getCol(), MissingCellPolicy.CREATE_NULL_AS_BLANK)
				.setCellValue("2");

		wb.getSheetAt(0).getRow(crc1.getRow()).getCell(crc1.getCol(), MissingCellPolicy.CREATE_NULL_AS_BLANK)
				.setCellValue("4");

		XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);

		// System.out.println(wb.getSheetAt(0).getRow(crc.getRow()).getCell(crc.getCol()).getRawValue());

		String rv = wb.getSheetAt(1).getRow(crc.getRow()).getCell(crc.getCol()).getRawValue();

		System.out.println(rv);

		return System.currentTimeMillis() - l;
	}
}
