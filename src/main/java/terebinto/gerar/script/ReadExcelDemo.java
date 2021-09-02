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

        String diretorio = "/tmp/notas/";
        String nomeArquivo = "NF_679";

        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        String date = new SimpleDateFormat("ddMMyyyyss").format(timestamp.getTime());
        ArrayList<Long> listaId = new ArrayList<>();

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
                Number sid = new Long(0);

                if (row.getRowNum() > 0) {

                    int cont = 0;


                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                if (cont == 0) {
                                    sid = cell.getNumericCellValue();
                                    listaId.add(sid.longValue());
                                } else if (cont == 1) {
                                    obuild = cell.getNumericCellValue();
                                }
                                cont++;
                                break;
                            case Cell.CELL_TYPE_STRING:
                                if (cont == 0) {
                                    sid = Long.valueOf(cell.getStringCellValue());
                                    listaId.add(sid.longValue());
                                } else if (cont == 1) {
                                    obuild = Long.valueOf(cell.getStringCellValue());
                                }
                                cont++;
                                break;
                        }
                    }

                    if (obuild.longValue() > 0) {
                        String script = "update compras.identificador set obuid ='" + obuild.longValue() + "', id_cliente= 3270840 Where id = " + sid.longValue() + " and obuid is null;" + "\n";
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

        for (Long serial : listaId) {

            if (cont2 == listaId.size()) {
                gravarArqQA.printf(String.valueOf(serial.longValue()));
            } else {
                gravarArqQA.printf(String.valueOf(serial.longValue()).concat(","));

            }

            cont2++;
        }

        gravarArqQA.printf(") and obuid is NULL");
        gravarArqQA.close();

        //gerar arquivo de qa
        FileWriter arqTeste = new FileWriter(diretorio + nomeArquivo.concat("-UPDATE-").concat(date) + ".sql");
        PrintWriter gravarArqTeste = new PrintWriter(arqTeste);

        String scriptUP = "select count(*) from compras.identificador i where id in ( ";

        gravarArqTeste.printf(scriptUP);

        int cont3 = 1;

        for (Long serial : listaId) {

            if (cont3 == listaId.size()) {
                gravarArqTeste.printf(String.valueOf(serial.longValue()));
            } else {
                gravarArqTeste.printf(String.valueOf(serial.longValue()).concat(","));

            }

            cont3++;
        }


        gravarArqTeste.printf(") and obuid is not NULL");
        gravarArqTeste.close();

        System.out.println("Total de registros encontrados na nota" + nomeArquivo.concat(".xlsx") + ".......: " + listaId.size());

    }
}

