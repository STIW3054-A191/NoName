package groupasg;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

import java.util.*;


public class ReportAsg {
    int count = 0;
    String name, matricno;
    int wmc, dit, noc, cbo, rfc, lcom;
    ArrayList<ReportAsg> list = new ArrayList<ReportAsg>( );

    ReportAsg(String matricno, String name, int wmc, int dit, int noc, int cbo, int rfc, int lcom) throws IOException {
        this.matricno = matricno;
        this.name = name;
        this.wmc = wmc;
        this.dit =dit;
        this.noc = noc;
        this.cbo = cbo;
        this.rfc = rfc;
        this.lcom =lcom;
    }

    ReportAsg(int count, String matricno, String name) throws IOException {
        this.count = count;
        this.matricno = matricno;
        this.name = name;
    }

    ReportAsg() throws IOException {
    }

    public String getName() { return name; }

    public void setName(String newName) {
        this.name = newName;
    }

    public String getMatricNo() {
        return matricno;
    }

    public void setMatricno(String matricno) {
        this.matricno = matricno;
    }

    public int getWmc() { return wmc; }

    public void setWmc(int wmc) { this.wmc = wmc; }

    public int getDit() { return dit; }

    public void setDit(int dit) { this.dit = dit; }

    public int getNoc() { return noc; }

    public void setNoc(int noc) { this.noc = noc; }

    public int getCbo() { return cbo; }

    public void setCbo(int cbo) { this.cbo = cbo; }

    public int getRfc() { return rfc; }

    public void setRfc(int rfc) { this.rfc = rfc; }

    public int getLcom() { return lcom; }

    public void setLcom(int lcom) { this.lcom = lcom; }

    public void createAccount(String matricno, String name, int wmc, int dit, int noc, int cbo, int rfc, int lcom) throws IOException {
        ReportAsg s = new ReportAsg(matricno, name, wmc, dit, noc, cbo, rfc, lcom);
        int c =0;
        list.add(s);
        try {
            XSSFWorkbook workbook = new XSSFWorkbook( );
            XSSFSheet sheet = workbook.createSheet("Report of Assignment 1 ");// creating a blank sheet

            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);
            Cell cell5 = row.createCell(5);
            Cell cell6 = row.createCell(6);
            Cell cell7 = row.createCell(7);
            Cell cell8 = row.createCell(8);
            cell.setCellValue("No");
            cell1.setCellValue("Matric No.");
            cell2.setCellValue("Name");
            cell3.setCellValue("WMC");
            cell4.setCellValue("DIT");
            cell5.setCellValue("NOC");
            cell6.setCellValue("CBO");
            cell7.setCellValue("RFC");
            cell8.setCellValue("LCOM");

            int rownum = 1;
            for (ReportAsg listf : list) {
                Row row1 = sheet.createRow(rownum++);
               //createList(c,listf, row1);
                //c++;
            }
            FileOutputStream out = new FileOutputStream(new File("C:\\Users\\HALIMAH\\IdeaProjects\\GroupAsg\\src\\Excel\\ReportASG.xlsx")); // file name with path
            for(int j =0; j<3;j++){
                sheet.autoSizeColumn(j);
            }
            workbook.write(out);
            out.close( );

        } catch (Exception e) {
            e.printStackTrace( );
        }
    }


    public static void createList(int c,ReportAsg f, Row row) // creating cells for each row

    {
        Cell cell = row.createCell(0);
        Cell cell1 = row.createCell(1);
        Cell cell2 = row.createCell(2);
        Cell cell3 = row.createCell(3);
        Cell cell4 = row.createCell(4);
        Cell cell5 = row.createCell(5);
        Cell cell6 = row.createCell(6);
        Cell cell7 = row.createCell(7);
        Cell cell8 = row.createCell(8);

        cell.setCellValue(c);
        cell1.setCellValue(f.getMatricNo( ));
        cell2.setCellValue(f.getName( ));
        cell3.setCellValue(f.getWmc());
        cell4.setCellValue(f.getDit( ));
        cell5.setCellValue(f.getNoc( ));
        cell6.setCellValue(f.getCbo( ));
        cell7.setCellValue(f.getRfc( ));
        cell8.setCellValue(f.getLcom( ));
    }

    public void StudentList() {
        for (ReportAsg std : list) {
            System.out.println(std.getMatricNo( ) + " " + std.getName( ));
        }
    }

}

