package org.wesol.app;

import org.apache.commons.cli.*;
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
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;

public class MainApp {
    static String baseFilepath = "";
    static String driverPath = "";
    static String excelFileName = "";

    static Set<String> uniqueSet = new HashSet<>();
    static WebDriver driver;
    static String downloadFilePath = "";
    static String excelPath = "";
    static String folderAddressExcel = "";
    static final String baseURL = "https://bocaodientu.dkkd.gov.vn/egazette/Forms/Egazette/DefaultAnnouncements.aspx";
    static SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");

    public static void main(String[] args) {
        //parse option
        parseOption(args);
        prepare();
        doCrawInfo();
        System.out.println("Download info company completed!");
    }

    private static void parseOption(String[] args) {
        Options options = new Options();

        Option input = new Option("d", "driver", true, "driver file path");
        input.setRequired(true);
        options.addOption(input);

        Option output = new Option("o", "output", true, "output folder");
        output.setRequired(true);
        options.addOption(output);

        Option excelFile = new Option("e", "excel", true, "excel file name");
        output.setRequired(true);
        options.addOption(excelFile);

        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        CommandLine cmd;

        try {
            cmd = parser.parse(options, args);
            driverPath = cmd.getOptionValue("driver");
            baseFilepath = cmd.getOptionValue("output");
            excelFileName = cmd.getOptionValue("excel");
            excelPath = baseFilepath;

            System.out.println(driverPath);
            System.out.println(baseFilepath);
            System.out.println(excelFileName);
        } catch (ParseException e) {
            System.out.println(e.getMessage());
            formatter.printHelp("CrawCompanyInfo", options);

            System.exit(1);
        }
    }

    public static void prepare() {
        //new folder:
        String folderName = sdf.format(new Date());
        System.out.println("create folder " + folderName);
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
            System.out.println("backing up excel file...");
            Files.copy(new File(excelPath + File.separator + excelFileName).toPath(),
                    new File(excelPath + File.separator + folderName + "_bk_" + excelFileName).toPath());
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static List<CompanyInfo> doCrawInfo() {
        //setup chrome driver
        System.setProperty("webdriver.chrome.driver", driverPath);
        HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
        chromePrefs.put("profile.default_content_settings.popups", 0);
        chromePrefs.put("download.default_directory", downloadFilePath);
        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("prefs", chromePrefs);
        DesiredCapabilities cap = DesiredCapabilities.chrome();
        cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
        cap.setCapability(ChromeOptions.CAPABILITY, options);
        driver = new ChromeDriver(cap);

        // call GET page
        driver.get(baseURL);
        int currentPage = 1;


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
                link.setAddress(folderAddressExcel + "/" + companyInfo.getMsdn() + ".pdf");
                cell.setCellValue(companyInfo.getMsdn() + ".pdf");
                cell.setHyperlink(link);
            }


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
        System.out.println("load all downloaded file name...");

        try (Stream<Path> paths = Files.walk(Paths.get(baseFilepath))) {
            paths.filter(Files::isRegularFile)
                    .filter(t -> t.getFileName().toString().endsWith("pdf"))
                    .forEach(file -> uniqueSet.add(file.getFileName().toString().replace(".pdf", "")));
        } catch (Exception e) {
            System.out.println("Load files failed");
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
                break;
            }
        }
        return success;
    }

    private static void waitForFileDownload() {
        FluentWait<WebDriver> wait = new FluentWait(driver)
                .withTimeout(5000, TimeUnit.MILLISECONDS)
                .pollingEvery(200, TimeUnit.MILLISECONDS);

        File f = new File(baseFilepath + "/" + "new_announcement.pdf");
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
        return companyInfos;
    }

}
