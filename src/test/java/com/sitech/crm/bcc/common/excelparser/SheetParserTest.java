package com.sitech.crm.bcc.common.excelparser;

import com.sitech.crm.bcc.common.domain.*;

import com.sitech.crm.bcc.common.excelparser.exception.ExcelParsingException;
import com.sitech.crm.bcc.common.excelparser.interfaces.ParserExceptionConsumer;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static java.math.BigDecimal.ROUND_FLOOR;
import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertThat;

public class SheetParserTest {

    InputStream inputStream;

    @After
    public void tearDown() throws IOException {
        inputStream.close();
    }

    @Test
    public void shouldCreateEntityBasedOnAnnotationFromExcel97File() throws Exception {
        performTestUsing(openSheet("Student Profile.xls"));
    }

    @Test
    public void shouldCreateEntityBasedOnAnnotationFromExcel97File2() throws Exception {
        Workbook workBook = openWorkbook("Student Profile.xls");
        performTestUsing(workBook.getSheet("Sheet1"));
        SheetParser parser = new SheetParser();
        List<Section> entityList = parser.createEntity(workBook.getSheet("Sheet1"), Section.class, new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });

        Workbook workBook2 = openWorkbook("Student Profilecopy.xls");

        parser.createSheet(workBook2.getSheet("Sheet1"),Section.class,entityList,new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });

        writeWorkbook("Student Profile Result.xls",workBook2);

    }

    @Test
    public void shouldCreateEntityBasedOnAnnotationFromExcel2007File() throws Exception {
        performTestUsing(openSheet("Student Profile.xlsx"));
    }

    @Test
    public void shouldCreateEntityBasedOnAnnotationFromExcel2007File2() throws Exception {
        Workbook workBook = openWorkbook("Student Profile.xlsx");
        performTestUsing(workBook.getSheet("Sheet1"));
        SheetParser parser = new SheetParser();
        List<Section> entityList = parser.createEntity(workBook.getSheet("Sheet1"), Section.class, new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });

        Workbook workBook2 = openWorkbook("Student Profilecopy.xlsx");

        parser.createSheet(workBook2.getSheet("Sheet1"),Section.class,entityList,new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });

        writeWorkbook("Student Profile Result.xlsx",workBook2);
    }


    @Test
    public void shouldCallErrorHandlerWhenRowCannotBeParsed() throws Exception {
        final List<ExcelParsingException> errors = new ArrayList<ExcelParsingException>();
        SheetParser parser = new SheetParser();

        List<Section> entityList = parser.createEntity(openSheet("Errors.xlsx"), Section.class, new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                errors.add(e);
            }
        });

        assertThat(entityList.size(), is(1));
        Section section = entityList.get(0);
        assertThat(section.getStudents().get(0).getDateOfBirth(), is(nullValue(Date.class)));

        assertThat(errors.size(), is(3));
        assertThat(errors.get(0).getMessage(), is("Invalid date found in sheet Sheet1 at row 6, column 4"));
        assertThat(errors.get(1).getMessage(), is("Invalid date found in sheet Sheet1 at row 7, column 4"));
        assertThat(errors.get(2).getMessage(), is("Invalid date found in sheet Sheet1 at row 8, column 4"));
    }

    private void performTestUsing(Sheet sheet) {
        SheetParser parser = new SheetParser();
        List<Section> entityList = parser.createEntity(sheet, Section.class, new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });
        assertThat(entityList.size(), is(1));
        Section section = entityList.get(0);
        assertThat(section.getYear(), is("IV"));
        assertThat("B", section.getSection(), is("B"));
        assertThat(section.getStudents().size(), is(3));

        assertThat(section.getStudents().get(0).getRoleNumber(), is(2001L));
        assertThat(section.getStudents().get(0).getName(), is("Adam"));

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy");

        assertThat(simpleDateFormat.format(section.getStudents().get(0).getDateOfBirth()), is("01/01/2002"));
        assertThat(section.getStudents().get(0).getFatherName(), is("A"));
        assertThat("D", section.getStudents().get(0).getMotherName(), is("D"));
        assertThat("XYZ", section.getStudents().get(0).getAddress(), is("XYZ"));
        assertThat(section.getStudents().get(0).getTotalScore(), is(nullValue(BigDecimal.class)));


        assertThat(section.getStudents().get(1).getRoleNumber(), is(2002L));
        assertThat(section.getStudents().get(1).getName(), is("Even"));
        assertThat(simpleDateFormat.format(section.getStudents().get(1).getDateOfBirth()), is("05/01/2002"));
        assertThat(section.getStudents().get(1).getFatherName(), is("B"));
        assertThat("D", section.getStudents().get(1).getMotherName(), is("E"));
        assertThat("XYZ", section.getStudents().get(1).getAddress(), is("ABX"));
        assertThat(section.getStudents().get(1).getTotalScore().setScale(2, ROUND_FLOOR), is(new BigDecimal("450.35")));



    }

    @Test
    public void shouldParseRowOrCoulmnEnd() throws IOException {
        Sheet sheet = openSheet("Student Profile.xlsx");
        SheetParser parser = new SheetParser();

        int end = parser.getRowOrColumnEnd(sheet, Section.class);
        assertEquals(end, 2);

        int rowEnd = parser.getRowOrColumnEnd(sheet, Student.class);
        assertEquals(rowEnd, 8);
    }

    @Test
    public void shouldCalculateEnd() throws IOException {
        Sheet sheet = openSheet("Subjects.xlsx");
        SheetParser parser = new SheetParser();

        List<Subject> subjects = parser.createEntity(sheet, Subject.class, new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });
        Subject lastSubject = subjects.get(subjects.size() - 1);

        assertThat(lastSubject.getCode(), is("GE-101"));
        assertThat(lastSubject.getName(), is("Geography"));
        assertThat(lastSubject.getVolume(), is(7));
    }

    @Test
    public void shouldParserGroupShort() throws IOException {
        Sheet sheet = openSheet("ty.xls");
        SheetParser parser = new SheetParser();

        final List<ExcelParsingException> excelInvalidCells = new ArrayList<ExcelParsingException>();

        ParserExceptionConsumer<ExcelParsingException> errorHandler = new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                excelInvalidCells.add(e);
            }
        };

        List<GroupShort> groupShorts = parser.createEntity(sheet, GroupShort.class, errorHandler);

        GroupShort lastgroupShort = groupShorts.get(groupShorts.size() - 1);

        assertThat(lastgroupShort.getMemtelNo(), is("18617392774"));
        assertThat(lastgroupShort.getShortNo(), is("92774"));
        assertThat(lastgroupShort.getFav(), is("套餐1.00元/100分钟|gl.grpmem.vpmn.170"));
        assertThat(lastgroupShort.getEffectDate(), is("立即生效"));
    }

    @Test
    public void shouldParserGroupColorBell() throws IOException {
        Sheet sheet = openSheet("pg.vo.gpcr.dwzf.xls");
        SheetParser parser = new SheetParser();

        final List<ExcelParsingException> excelInvalidCells = new ArrayList<ExcelParsingException>();

        ParserExceptionConsumer<ExcelParsingException> errorHandler = new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                excelInvalidCells.add(e);
            }
        };

        List<GroupColorBell> groupColorBells = parser.createEntity(sheet, GroupColorBell.class, errorHandler);

        GroupColorBell lastgroupColorBell = groupColorBells.get(groupColorBells.size() - 1);

        assertThat(lastgroupColorBell.getMemtelNo(), is("18617392775"));
        assertThat(lastgroupColorBell.getFav(), is("集团彩铃（单位付费）减免个人彩铃费|gl_jtcl_dwzffee"));
        assertThat(lastgroupColorBell.getEffectDate(), is("立即生效"));
    }

    @Test
    public void shouldParserGroupAccountMember() throws IOException {
        Sheet sheet = openSheet("acctAddMemTemplate.xlsx");
        SheetParser parser = new SheetParser();

        final List<ExcelParsingException> excelInvalidCells = new ArrayList<ExcelParsingException>();

        ParserExceptionConsumer<ExcelParsingException> errorHandler = new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                excelInvalidCells.add(e);
            }
        };

        List<GroupAccountMember> groupAccountMembers = parser.createEntity(sheet, GroupAccountMember.class, errorHandler);

        GroupAccountMember lastgroupAccountMember = groupAccountMembers.get(groupAccountMembers.size() - 1);

        assertThat(lastgroupAccountMember.getMemtelNo(), is("18617392776"));
        assertThat(lastgroupAccountMember.getPayType(), is("限额付费"));
        assertThat(lastgroupAccountMember.getPayexpenseType(), is("节约归己"));
        assertThat(lastgroupAccountMember.getIsTransBill(), is("否"));
        assertThat(lastgroupAccountMember.getEffectType(), is("当月生效"));
        assertThat(lastgroupAccountMember.getMaxQuota(), is(5d));
    }

    private Sheet openSheet(String fileName) throws IOException {
        inputStream = getClass().getClassLoader().getResourceAsStream(fileName);
        Workbook workbook;
        if(fileName.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook.getSheet("Sheet1");
    }

    private Workbook openWorkbook(String fileName) throws IOException {
        inputStream = getClass().getClassLoader().getResourceAsStream(fileName);
        Workbook workbook;
        if(fileName.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    private void writeWorkbook(String fileName,Workbook workbook) throws IOException {
        FileOutputStream excelFileOutPutStream = new FileOutputStream(fileName);//写数据到这个路径上
        workbook.write(excelFileOutPutStream);
        excelFileOutPutStream.flush();
        excelFileOutPutStream.close();
    }

}
