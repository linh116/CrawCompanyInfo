package org.wesol.app;

import org.apache.commons.cli.*;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.wesol.entity.CompanyInfo;
import org.wesol.helper.ClientOptionHelper;
import org.wesol.helper.LinkHelper;
import org.wesol.helper.SeleniumHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;

import static org.wesol.constant.Constant.BASE_URL;

public class MainApp {
    //global input
    static String baseFilepath = "";
    static String driverPath = "";
    static String excelFileName = "";
    static WebDriver driver;

    //global
    static Set<String> uniqueSet = new HashSet<>();
    static String downloadFilePath = "";
    static String excelPath = "";
    static String folderAddressExcel = "";
    static int stopPageNum = 0;
    static SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");

    public static void main(String[] args) {
        //parse option
        parseOption(args);
        prepare();
        doCrawInfo();
        System.out.println("Download info company completed!");
    }

    private static void parseOption(String[] args) {
        CommandLine cmd = ClientOptionHelper.createCommandLineOption(args);
        driverPath = cmd.getOptionValue("driver");
        baseFilepath = cmd.getOptionValue("output");
        excelFileName = cmd.getOptionValue("excel");
        String stopPageStr = cmd.getOptionValue("stoppage");
        if (stopPageStr != null) {
            stopPageNum = Integer.valueOf(stopPageStr);
        }

        excelPath = baseFilepath;
    }

    public static void prepare() {
        //new folder:
        String folderName = sdf.format(new Date());
        System.out.println("Create folder " + folderName);
        folderAddressExcel = folderName;
        File newFolder = new File(baseFilepath + File.separator + folderName);
        if (!newFolder.exists()) {
            newFolder.mkdirs();
        }
        downloadFilePath = newFolder.getPath();

        // load file name to avoid duplicate download
        loadFileName();

        //backup excel file
        try {
            System.out.println("Backing up excel file...");
            Files.copy(new File(excelPath + File.separator + excelFileName).toPath(),
                    new File(excelPath + File.separator + folderName + "_bk_" + excelFileName).toPath());
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static List<CompanyInfo> doCrawInfo() {
        //setup chrome driver
        driver = SeleniumHelper.createDriver(driverPath, downloadFilePath);

        // call GET page
        driver.get(BASE_URL);
        driver.manage().window().maximize();
        int currentPage = 1;

        List<CompanyInfo> ret = new ArrayList<>();
        try {
            while (true) { //crawl till "stopPage" continue pages do nothing
                System.out.println("----- Loading page: " + currentPage + " -----");

                //ensure webpage loaded
                WebDriverWait waiter = new WebDriverWait(driver, 10);
                waiter.until(d -> d.findElement(By.tagName("table")));
                WebElement table = driver.findElement(By.tagName("table"));

                //get info and download
                List<CompanyInfo> companyInfos = getInfoCompany(table);
                if (stopPageNum != 0 && currentPage >= stopPageNum) {
                    return ret;
                }

                //write excel
                if (companyInfos.size() > 0) {
                    ret.addAll(companyInfos);
                    writeToExcel(companyInfos);
                }

                if (currentPage == 25) { // stop when page is max
                    return ret;
                }

                //go to next page
                WebElement pager = table.findElement(By.className("Pager"));
                List<WebElement> pages = pager.findElement(By.tagName("td")).findElements(By.tagName("td"));
                goToNextPage(pages, currentPage);
                currentPage++;
            }
        } catch (Exception e) {
            e.printStackTrace();
            return ret;
        }
    }


    synchronized private static List<CompanyInfo> getInfoCompany(WebElement table) {
        List<WebElement> rows = table.findElements(By.tagName("tr"));
        List<WebElement> cols;
        String companyName, msdn, province;
        List<CompanyInfo> companyInfos = new ArrayList<>();
        int downloadedFileName = 0;
        for (int i = 1; i < rows.size() - 2; i++) { //skip header
            cols = rows.get(i).findElements(By.tagName("td"));
            try {
                companyName = cols.get(1).findElement(By.tagName("p")).getText();
                msdn = cols.get(1).findElement(By.tagName("div")).getText().replace("MÃ SỐ DN: ", "");
                province = cols.get(2).getText();
                if (uniqueSet.contains(msdn)) {
                    continue;
                } else {
                    uniqueSet.add(msdn);
                }
                System.out.println(String.format("\t %s::%s::%s", companyName, msdn, province));

                //click button download
                cols.get(4).findElement(By.tagName("input")).click();

                //add to map rename later
                String fileName = downloadedFileName == 0 ? "new_announcement.pdf" : "new_announcement (" + (downloadedFileName) + ").pdf";
                downloadedFileName++;

                CompanyInfo company = new CompanyInfo(msdn, companyName, province);
                company.setFileName(fileName);
                companyInfos.add(company);

            } catch (NoSuchElementException e) {
                System.err.println("not found ele " + i);
            }
        }
        boolean success;
        try {
            Thread.sleep(5000);
            changeFileName(companyInfos);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        return companyInfos;
    }


    private static void writeToExcel(List<CompanyInfo> companyInfos) {
        System.out.println("write to excel " + companyInfos.size() + " files");
        String filePath = excelPath + File.separator + excelFileName;

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = null;
        int rowCount = 0;
        try {
            File file = new File(filePath);
            if (!file.exists()) {
                file.createNewFile();
            }
            workbook = WorkbookFactory.create(new File(filePath));
            CellStyle styleHyperLink = LinkHelper.getStyleHyperLink(workbook);
            CellStyle styleLock = LinkHelper.getStyleLocked(workbook);
            // Getting the Sheet at index zero
            Sheet sheet = workbook.getSheetAt(0);
            int column;
            Row row;
            Cell cell;
            boolean hasHeader = sheet.getRow(0) != null;
            if (hasHeader == false) { //create header in case header not exist
                row = sheet.createRow(0);
                column = -1;
                cell = row.createCell(++column);
                cell.setCellValue("Time Craw");
                cell = row.createCell(++column);
                cell.setCellValue("MSDN");
                cell = row.createCell(++column);
                cell.setCellValue("Tên Công Ty");
                cell = row.createCell(++column);
                cell.setCellValue("Tỉnh/TP");
                cell = row.createCell(++column);
                cell.setCellValue("Link PDF");
            }
            for (CompanyInfo companyInfo : companyInfos) {
                row = sheet.getRow(1);

                // If the row exist in destination, push down all rows by 1 else create a new row
                if (row != null) {
                    sheet.shiftRows(1, sheet.getLastRowNum(), 1);
                }

                //fill value
                row = sheet.createRow(1);

                column = -1;
                cell = row.createCell(++column);
                cell.setCellValue(new Date());
                cell.setCellStyle(styleLock);

                cell = row.createCell(++column);
                cell.setCellValue(companyInfo.getMsdn());
                cell.setCellStyle(styleLock);

                cell = row.createCell(++column);
                cell.setCellValue(companyInfo.getCompanyName());
                cell.setCellStyle(styleLock);

                cell = row.createCell(++column);
                cell.setCellValue(companyInfo.getProvince());
                cell.setCellStyle(styleLock);

                cell = row.createCell(++column);
                XSSFHyperlink link = LinkHelper.createHyperlink(HyperlinkType.FILE);
                link.setAddress(folderAddressExcel + "/" + companyInfo.getMsdn() + ".pdf");
                cell.setCellValue(companyInfo.getMsdn() + ".pdf");
                cell.setHyperlink(link);
                cell.setCellStyle(styleHyperLink);
            }

            //write file
            FileOutputStream outputStream = new FileOutputStream(excelPath + File.separator + "infos-temp.xlsx");
            workbook.write(outputStream);

            // Closing the workbook
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void loadFileName() {
        //load all file
        System.out.print("Load all downloaded file name... ");

        try (Stream<Path> paths = Files.walk(Paths.get(baseFilepath))) {
            paths.filter(Files::isRegularFile)
                    .filter(t -> t.getFileName().toString().endsWith("pdf"))
                    .forEach(file -> uniqueSet.add(file.getFileName().toString().replace(".pdf", "")));
        } catch (Exception e) {
            System.out.println("Load files failed");
        }

        System.out.println("loaded " + uniqueSet.size() + " file(s).");
    }

    private static void goToNextPage(List<WebElement> pages, int currentPage) {
        if (currentPage < 6) {
            pages.get(currentPage).click();
        } else {
            pages.get(currentPage % 6 + 2).click();
        }
    }

    private static void changeFileName(List<CompanyInfo> companyInfos) {
        List<CompanyInfo> failed = new ArrayList<>();
        for (CompanyInfo companyInfo : companyInfos) {
            if (companyInfo.getFileName() == null || companyInfo.getFileName().trim().equals("")) {
                System.out.println(String.format("Company %s doesnt have any file name", companyInfo.getMsdn()));
                continue;
            }
            if (changeFileName(companyInfo.getFileName(), companyInfo.getMsdn()) == false) {
                failed.add(companyInfo);
            }
        }

        companyInfos.removeAll(failed);

    }

    private static boolean changeFileName(String oldFileName, String newFileName) {
        File file = new File(downloadFilePath + "/" + oldFileName);


        boolean success = file.renameTo(new File(downloadFilePath + "/" + newFileName + ".pdf"));
        int failed = 0;
        while (!success) {
            System.out.println("failed to rename " + oldFileName + "...");
            try {
                Thread.sleep(200);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            success = file.renameTo(new File(downloadFilePath + "/" + newFileName + ".pdf"));
            if (success) {
                System.out.println("Retry success.");
                break;
            } else {
                failed++;
            }
            if (failed == 5) {
                System.out.println("Stop rename " + oldFileName + " to " + newFileName + " because failed 5 times");
                return false;
            }
        }
        return success;
    }

}
