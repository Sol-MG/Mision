package pe.uni.sistemas.horario;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;

/**
 * Created by USUARIO on 31/05/2015.
 */
public class HorarioVuelo {
    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook();

        try {
            FileOutputStream fileOut = new FileOutputStream("D:\\ExcelHorario\\Myexcel.xls");
            Sheet hoja1 = wb.createSheet("Hoja1");
            Row fila0 = hoja1.createRow(1);
            Cell cell = fila0.createCell(8);
            HSSFRichTextString texto = new HSSFRichTextString("Horario  de  Salida  y  Entrada del  Aeropuesto  X");
            cell.setCellValue(texto);


            Row fila = hoja1.createRow(3);
            for (int i=1; i<3;i++){
                Cell celli = fila.createCell(6*i);
                celli.setCellValue("Fecha");
                Cell cell2i = fila.createCell(7);
                cell2i.setCellValue("Destino");
                Cell cell2 = fila.createCell(13);
                cell2.setCellValue("Origen");
                Cell cell3i = fila.createCell(8);
                cell3i.setCellValue("Salida");
                Cell cell4 = fila.createCell(14);
                cell4.setCellValue("Entrada");
            }

            for (int i=0;i<3;i++) {
                Row fila1 = hoja1.createRow(4 + i);
                for (int j=0;j<2;j++) {
                    Cell cell0 = fila1.createCell(6+6*j);
                    cell0.setCellValue("01-"+(10+2*j)+"-15");
                    Cell cell2 = fila1.createCell(7+6*j);
                    cell2.setCellValue("Miami");
                    Cell cell1 = fila1.createCell(8+6*j);
                    cell1.setCellValue((7+(j+1)*(1+i)) +":"+(10+ j*(10+2*i))+" AM");
                }
            }
            for (int i=0;i<3;i++) {
                Row fila2 = hoja1.createRow(7+i);
                for (int j=0;j<2;j++) {
                    Cell cell0 = fila2.createCell(6+6*j);
                    cell0.setCellValue("02-"+(4*(i+1)+2*j)+"-15");
                    Cell cell2 = fila2.createCell(7+6*j);
                    cell2.setCellValue("Mexico");
                    Cell cell1 = fila2.createCell(8+6*j);
                    cell1.setCellValue(1+j +":"+(10+ j*(10+2*i))+" PM");
                }
            }

            wb.write(fileOut);
            fileOut.close();
        }
        catch (Exception e){

        }
    }
}
