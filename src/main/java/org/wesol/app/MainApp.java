package org.wesol.app;

import org.apache.commons.io.FileUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.wesol.entity.CompanyInfo;
import org.wesol.helper.LinkHelper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.TimeUnit;

public class MainApp {
    static Set<String> uniqueSet = new HashSet<>();
    static String downloadFilepath = "D:\\Coding\\Code\\data";
    static WebDriver driver;
    static String baseURL = "https://bocaodientu.dkkd.gov.vn/egazette/Forms/Egazette/DefaultAnnouncements.aspx";
    static String driverPath = "D:\\Coding\\Code\\chromedriver\\chromedriver.exe";
    static String excelPath = downloadFilepath;

    public static void main(String[] args) {
        //prepare
        prepare();
        List<CompanyInfo> companyInfos = doCrawInfo();
        System.out.println("Download info company completed!");
    }

    public static void prepare() {
        File directory = new File(downloadFilepath);
        File[] files = directory.listFiles(File::isFile);

        if (files != null) {
            for (File file : files) {
                if (file.getName().startsWith("new_announcement")) {
                    try {
                        FileUtils.forceDelete(file);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
            System.out.println("Loaded " + uniqueSet.size() + " files");
        }
    }

    public static List<CompanyInfo> doCrawInfo() {
        //setup chrome driver
        System.setProperty("webdriver.chrome.driver", driverPath);
        HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
        chromePrefs.put("profile.default_content_settings.popups", 0);
        chromePrefs.put("download.default_directory", downloadFilepath);
        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("prefs", chromePrefs);
        DesiredCapabilities cap = DesiredCapabilities.chrome();
        cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
        cap.setCapability(ChromeOptions.CAPABILITY, options);
        driver = new ChromeDriver(cap);

        // call GET page
        driver.get(baseURL);
        int currentPage = 1;

        // load file name to avoid duplicate download
        loadFileName();
        int numPageDoNothing = 0;
        List<CompanyInfo> ret = new ArrayList<>();
        try {
            while (true) { //crawl till 5 continue pages do nothing
                System.out.println("loading page " + currentPage);
                WebDriverWait waiter = new WebDriverWait(driver, 10);
                waiter.until(d -> d.findElement(By.tagName("table")));
                WebElement table = driver.findElement(By.tagName("table"));

                //get info and download
                List<CompanyInfo> companyInfos = getInfoCompany(table);
                numPageDoNothing = companyInfos.size() > 0 ? 0 : numPageDoNothing + 1;
                if (numPageDoNothing >= 5) {
                    return ret;
                }

                //write excel
                if (companyInfos.size() > 0) {
                    ret.addAll(companyInfos);
                    writeToExcel(companyInfos);
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

    private static void writeToExcel(List<CompanyInfo> companyInfos) {
        System.out.println("write to excel " + companyInfos.size() + " files");
        String filePath = excelPath + "\\infos.xlsx";

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = null;
        int rowCount = 0;
        try {
            File file = new File(filePath);
            if (!file.exists()) {
                file.createNewFile();
            }
            workbook = WorkbookFactory.create(new File(filePath));

            // Getting the Sheet at index zero
            Sheet sheet = workbook.getSheetAt(0);
            int column;
            Row row;
            Cell cell;
            boolean hasHeader = sheet.getRow(0) != null;
            if (hasHeader == false) { //create header
                row = sheet.createRow(0);
                column = -1;
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

                row = sheet.createRow(1);


                column = -1;
                cell = row.createCell(++column);
                cell.setCellValue(companyInfo.getMsdn());
                cell = row.createCell(++column);
                cell.setCellValue(companyInfo.getCompanyName());
                cell = row.createCell(++column);
                cell.setCellValue(companyInfo.getProvince());
                cell = row.createCell(++column);
                XSSFHyperlink link = LinkHelper.createHyperlink(HyperlinkType.FILE);
                link.setAddress(companyInfo.getMsdn() + ".pdf");
                cell.setCellValue(companyInfo.getMsdn() + ".pdf");
                cell.setHyperlink(link);
            }


            FileOutputStream outputStream = new FileOutputStream(excelPath + "\\infos-temp.xlsx");
            workbook.write(outputStream);
            // Closing the workbook
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void loadFileName() {
        File directory = new File(downloadFilepath);
        File[] files = directory.listFiles(File::isFile);

        if (files != null) {
            for (File file : files) {
                uniqueSet.add(file.getName().replace(".pdf", ""));
            }
            System.out.println("Loaded " + uniqueSet.size() + " files");
        }
    }

    private static void goToNextPage(List<WebElement> pages, int currentPage) {
        if (currentPage < 6) {
            pages.get(currentPage).click();
        } else {
            pages.get(currentPage % 6 + 2).click();
        }
    }

    private static void changeFileName(List<CompanyInfo> companyInfos) {
        for (CompanyInfo companyInfo : companyInfos) {
            if (companyInfo.getFileName() == null || companyInfo.getFileName().trim().equals("")) {
                System.out.println(String.format("Company %s doesnt have any file name", companyInfo.getMsdn()));
                continue;
            }
            changeFileName(companyInfo.getFileName(), companyInfo.getMsdn());
        }
    }

    private static boolean changeFileName(String oldFileName, String newFileName) {
        File file = new File(downloadFilepath + "/" + oldFileName);


        boolean success = file.renameTo(new File(downloadFilepath + "/" + newFileName + ".pdf"));
        int failed = 0;
        while (!success) {
            System.out.println("failed to rename " + oldFileName + "...");
            try {
                Thread.sleep(200);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            success = file.renameTo(new File(downloadFilepath + "/" + newFileName + ".pdf"));
            if (success) {
                System.out.println("Retry success.");
                break;
            } else {
                failed++;
            }
            if (failed == 5) {
                System.out.println("Stop rename " + oldFileName + " to " + newFileName + " because failed 5 times");
                break;
            }
        }
        return success;
    }

    private static void waitForFileDownload() {
        FluentWait<WebDriver> wait = new FluentWait(driver)
                .withTimeout(5000, TimeUnit.MILLISECONDS)
                .pollingEvery(200, TimeUnit.MILLISECONDS);

        File f = new File(downloadFilepath + "/" + "new_announcement.pdf");
        wait.until((WebDriver wd) -> f.exists());
    }

    private static File getLastModified(String directoryFilePath) {
        File directory = new File(directoryFilePath);
        File[] files = directory.listFiles(File::isFile);
        long lastModifiedTime = Long.MIN_VALUE;
        File chosenFile = null;

        if (files != null) {
            for (File file : files) {
                if (file.lastModified() > lastModifiedTime) {
                    chosenFile = file;
                    lastModifiedTime = file.lastModified();
                }
            }
        }

        return chosenFile;
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
                System.out.println(String.format("%s::%s::%s", companyName, msdn, province));
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
//        for (Map.Entry<String,String> t : mapFileName.entrySet()){
//            success = changeFileName(t.getKey(), t.getValue());
//            while (success == false){
//                System.out.println("failed to change file name " + t.getKey());
//                success = changeFileName(t.getKey(), t.getValue());
//                if(success == true){
//                    System.out.println("Retry change name file " + t.getKey() + " success.");
//                }
//            }
//        }
        return companyInfos;
    }

}
