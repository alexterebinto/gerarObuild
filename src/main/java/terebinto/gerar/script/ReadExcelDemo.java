package terebinto.gerar.script;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;


public class ReadExcelDemo {
    public static void main(String[] args) throws IOException {
        ArrayList<String> id = new ArrayList<>();
        ArrayList<String> obiud = new ArrayList<>();
        int linhas = 0;

        String diretorio = "/tmp/";
        String nomeArquivo = "teste1";

        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        String date = new SimpleDateFormat("ddMMyyyyss").format(timestamp.getTime());
        ArrayList<String> listaId = new ArrayList<>();

        try {
            FileInputStream file = new FileInputStream(new File(diretorio + nomeArquivo + ".xlsx"));
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            Row header = sheet.getRow(0);
            int n = header.getLastCellNum();
            linhas = n + 1;
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            FileWriter arq = new FileWriter(diretorio + nomeArquivo.concat(date) + ".sql");
            PrintWriter gravarArq = new PrintWriter(arq);


            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                Number obuild = new Long(0);
                String sid = "";

                if (row.getRowNum() > 0) {

                    int cont = 0;


                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();


                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC:

                                if (cont == 1) {
                                    obuild = cell.getNumericCellValue();

                                }
                                cont++;
                                break;
                            case Cell.CELL_TYPE_STRING:

                                if (cont == 0) {
                                    sid = cell.getStringCellValue();
                                    listaId.add(sid);
                                }

                                cont++;
                                break;
                        }

                    }

                    if (obuild.longValue() > 0) {
                        String script = "update compras.identificador set obuid ='" + obuild.longValue() + "', id_cliente= 3270840 Where id = " + sid + " and obuid is null;" + "\n";
                        gravarArq.printf(script);
                    }
                }


            }
            file.close();
            arq.close();
            gravarArq.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        //gerar arquivo de qa
        FileWriter arqQA = new FileWriter(diretorio + nomeArquivo.concat("QA").concat(date) + ".sql");
        PrintWriter gravarArqQA = new PrintWriter(arqQA);

        String scriptQA = "select count(*) from compras.identificador i where id in ( ";

        gravarArqQA.printf(scriptQA);

        int cont2 = 1;

        for (String serial : listaId) {

            if (cont2 == listaId.size()) {
                gravarArqQA.printf(serial);
            } else {
                gravarArqQA.printf(serial.concat(","));

            }

            cont2++;
        }

        gravarArqQA.printf(") and obuid is NULL");
        gravarArqQA.close();


    }
}

