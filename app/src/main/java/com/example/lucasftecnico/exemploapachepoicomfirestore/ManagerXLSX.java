package com.example.lucasftecnico.exemploapachepoicomfirestore;

import com.example.lucasftecnico.exemploapachepoicomfirestore.styles.StylesXLSX;
import com.google.android.gms.tasks.OnSuccessListener;
import com.google.firebase.firestore.CollectionReference;
import com.google.firebase.firestore.FirebaseFirestore;
import com.google.firebase.firestore.QueryDocumentSnapshot;
import com.google.firebase.firestore.QuerySnapshot;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ManagerXLSX {
    private FirebaseFirestore firebaseFirestore = FirebaseFirestore.getInstance();
    private CollectionReference collectionReference = firebaseFirestore.collection("Enderecos");

    public void criar(final OutputStream outputStream) {

        collectionReference.get().addOnSuccessListener(new OnSuccessListener<QuerySnapshot>() {
            @Override
            public void onSuccess(QuerySnapshot queryDocumentSnapshots) {
                List<Endereco> enderecos = new ArrayList<>();
                XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
                XSSFSheet xssfSheet = xssfWorkbook.createSheet("Endereços");
                Map<String, CellStyle> styles = StylesXLSX.createStyles(xssfWorkbook, xssfSheet);
                String titulo = "Endereços SENAI";
                String[] subtitulos = {"Rua", "Número", "CEP", "Bairro", "Cidade", "Estado", "Pais"};

                for (QueryDocumentSnapshot documentSnapshot : queryDocumentSnapshots) {
                    Endereco endereco = documentSnapshot.toObject(Endereco.class);
                    enderecos.add(endereco);
                }

                //desliga e liga linhas de grade
                xssfSheet.setDisplayGridlines(false); // Grades
                xssfSheet.setPrintGridlines(false);
                xssfSheet.setFitToPage(true);
                xssfSheet.setHorizontallyCenter(true);
                PrintSetup printSetup = xssfSheet.getPrintSetup();
                printSetup.setLandscape(true);

                // Cria título
                Row rowTitulo = xssfSheet.createRow(0);
                for (int i = 1; i < subtitulos.length; i++) {
                    rowTitulo.createCell(i).setCellStyle(styles.get("titulo"));
                }
                Cell cellTitulo = rowTitulo.createCell(0);
                cellTitulo.setCellValue(titulo);
                cellTitulo.setCellStyle(styles.get("titulo"));
                xssfSheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$G$1"));

                // Cria subtítulo
                Row rowSubtitulo = xssfSheet.createRow(1);
                for (int i = 0; i < subtitulos.length; i++) {
                    xssfSheet.setColumnWidth(i, 15 * 256); // setColumnWidth(indice, largura)
                    Cell cell = rowSubtitulo.createCell(i);
                    cell.setCellValue(subtitulos[i]);
                    cell.setCellStyle(styles.get("subtitulo"));
                }

                int rownum = 2;
                for (Endereco endereco : enderecos) {
                    String styleCorpo;
                    if (rownum % 2 == 0) {
                        styleCorpo = "corpo_01";
                    } else {
                        styleCorpo = "corpo_02";
                    }

                    Row row = xssfSheet.createRow(rownum++);
                    int cellnum = 0;
                    Cell cellRua = row.createCell(cellnum++);
                    cellRua.setCellValue(endereco.getRua());
                    cellRua.setCellStyle(styles.get(styleCorpo));

                    Cell cellNumero = row.createCell(cellnum++);
                    cellNumero.setCellValue(endereco.getNumero());
                    cellNumero.setCellStyle(styles.get(styleCorpo));

                    Cell cellCEP = row.createCell(cellnum++);
                    cellCEP.setCellValue(endereco.getCEP());
                    cellCEP.setCellStyle(styles.get(styleCorpo));

                    Cell cellBairro = row.createCell(cellnum++);
                    cellBairro.setCellValue(endereco.getBairro());
                    cellBairro.setCellStyle(styles.get(styleCorpo));

                    Cell cellCidade = row.createCell(cellnum++);
                    cellCidade.setCellValue(endereco.getCidade());
                    cellCidade.setCellStyle(styles.get(styleCorpo));

                    Cell cellEstado = row.createCell(cellnum++);
                    cellEstado.setCellValue(endereco.getEstado());
                    cellEstado.setCellStyle(styles.get(styleCorpo));

                    Cell cellPais = row.createCell(cellnum++);
                    cellPais.setCellValue(endereco.getPais());
                    cellPais.setCellStyle(styles.get(styleCorpo));
                }

                try {
                    xssfWorkbook.write(outputStream);
                    outputStream.close();
                    System.out.println("Arquivo Excel criado com sucesso!");

                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                    System.out.println("Arquivo não encontrado!");
                } catch (IOException e) {
                    e.printStackTrace();
                    System.out.println("Erro na edição do arquivo!");
                }
            }
        });
    }
}