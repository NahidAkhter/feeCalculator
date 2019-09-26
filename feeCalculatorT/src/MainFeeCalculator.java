import java.io.File;

import com.sapient.feecalculator.Constant.FILETYPE;
import com.sapient.feecalculator.reader.ITransactionManager;
import com.sapient.feecalculator.reader.TrasactionReader;


public class MainFeeCalculator {
	public static void main(String[] args) {
		
		File transactionfile = new File(new File("").getAbsolutePath(),"resource/Input_java.xlsx");
		ITransactionManager tranction= TrasactionReader.getTrasactionReaderInstance().readFile(FILETYPE.EXCEL,transactionfile);		
		tranction.displayTransactionReport();	
		
		
	}
}
