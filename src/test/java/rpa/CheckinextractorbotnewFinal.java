package rpa;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementClickInterceptedException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Pdf;
import org.openqa.selenium.PrintsPage;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.print.PrintOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.ITestContext;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;

import objects.SeleniumUtils;
import objects.Utility;

public class CheckinextractorbotnewFinal {
    static SeleniumUtils sel;
    static Utility utility;
    static String projDirPath;
    static Logger logger = LogManager.getLogger(CheckinextractorbotnewFinal.class);
    public static ExcelTracker audit ;
    // ---- Config -----------------------------------------------------------
    private static final int    WINDOW_DAYS     	= 30;
    private static final int    PRUNE_DAYS   		= 45;   								// MUST be > WINDOW_DAYS
    private static final int    RESTART_EVERY 		= 70;  									// relaunch browser every N downloads
    private static final int    ROW_FIND_ATTEMPTS 	= 4;
    private static final Path   SEEN_FILE    		= Paths.get("extracted_ids.json");
    private static final DateTimeFormatter STAMP = DateTimeFormatter.ofPattern("yyyy-MM-dd", Locale.ENGLISH);

    private static final Path AUDIT_DIR = Paths.get("audit");
    private static final Path AUDIT_FILE = Paths.get( "checkin_audit_" + LocalDateTime.now().format(STAMP) + ".xlsx"); 	// PHI: protect
    private static final Path   DOWNLOAD_DIR 		= Paths.get("downloads");     	// PHI: protect
    private static final String FACILITY     		= "Afra Wound Care Associates LLC";

    // driver is mutable: swapped on each browser restart.
    private WebDriver driver;
    private WebDriverWait wait;
    public CheckinextractorbotnewFinal() { }

    public CheckinextractorbotnewFinal(WebDriver driver) {
        setDriver(driver);
    }

    private void setDriver(WebDriver d) {
        this.driver = d;
        this.wait = new WebDriverWait(d, Duration.ofSeconds(20));
    }

    // ---- Locators ---------------------------------------------------------
    private static final By PICKER     			= By.id("scheduler-filter-bar-datepicker");
    private static final By CAL_TITLE  			= By.cssSelector("#scheduler-filter-bar-datepicker .v-date-picker-header__value button");
    private static final By PREV_ARROW 			= By.cssSelector( "#scheduler-filter-bar-datepicker button[aria-label='Previous month']");
    private static final By NEXT_ARROW 			= By.cssSelector( "#scheduler-filter-bar-datepicker button[aria-label='Next month']");
    private static final By TODAY_LINK 			= By.id("scheduler-filter-today-button");
    private static final By LOADING_OVERLAY 	= By.cssSelector("div.overlay.overlay-front, div.overlay-fixed");
    private static final By PROGRESS_NOTE_LINK  = By.cssSelector("a.progress-note-link");
    private static final By NOTE_CONTENT 		= By.id("divHideContent");

    private static final By MODAL_BG     		= By.cssSelector("div.reveal-modal-bg");
    private static final By CONFIRM_EXIT 		= By.id("confirmYes");   // "Exit"

    private static final Pattern NOTE_MRN 		= Pattern.compile("Patient Number:\\s*([A-Za-z0-9-]+)");

    private static final DateTimeFormatter HDR 	= DateTimeFormatter.ofPattern("MMMM yyyy", Locale.ENGLISH);
    private static final DateTimeFormatter MDY 	= DateTimeFormatter.ofPattern("MM/dd/yyyy", Locale.ENGLISH);
    
    // drop down note date is non-padded M/d/yyyy, e.g. 6/29/2026
    private static final DateTimeFormatter DROPDOWN_DATE = DateTimeFormatter.ofPattern("M/d/yyyy", Locale.ENGLISH);

    // Date-matched drop down item (not "first item").
    private By dropdownItemForDate(LocalDate date) {
        String shown = date.format(DROPDOWN_DATE);
        return By.xpath(
            "//div[contains(@class,'v-list') and contains(@class,'v-list--dense')]" +
            "/div//a[contains(normalize-space(.),'" + shown + "')]");
    }

    // =======================================================================
    //  MAIN FLOW
    // =======================================================================
    public void run() throws IOException, InterruptedException {
        Files.createDirectories(DOWNLOAD_DIR);
        Files.createDirectories(AUDIT_DIR); 
        SeenStore    store 			= new SeenStore(SEEN_FILE);
        audit 						= new ExcelTracker(AUDIT_FILE);
        logger.info("Audit file for this run: " + AUDIT_FILE.toAbsolutePath());
        LocalDate appToday 			= readAppToday();
        List<LocalDate> window 		= new ArrayList<>();
        
        for (int i = WINDOW_DAYS - 1; i >= 0; i--) window.add(appToday.minusDays(i)); // oldest -> newest

        int downloaded = 0, duplicates = 0, errors = 0;
        int sinceRestart = 0;   // downloads since last browser relaunch

        try {
            for (LocalDate date : window) {
                List<Appointment> harvested;
                try {
                    dismissModalIfPresent();
                    selectDay(date);
                    waitForTableReload();
                    harvested = harvestWithRetry(date);
                } catch (Exception e) {
                    logger.error("Could not open/harvest day " + date + " ("
                            + e.getClass().getSimpleName() + "): " + e.getMessage());
                    errors++;
                    recoverToScheduler();
                    continue;
                }

                List<String> workList = new ArrayList<>();
                for (Appointment a : harvested) {
                    if (store.isNew(a.id)) workList.add(a.id);
                    else duplicates++;
                }
                int expected = workList.size();
                logger.info("Day " + date + ": " + harvested.size() + " checked, "
                        + expected + " new to download");

                int dayDownloaded = 0, dayErrors = 0;

                for (String key : workList) {
                    if (!store.isNew(key)) continue;

                    // ---- periodic browser restart to prevent renderer OOM ----
                    if (sinceRestart >= RESTART_EVERY) {
                        logger.info("Restarting browser after " + sinceRestart
                                + " downloads (memory hygiene)");
                        try { restartBrowser(); }
                        catch (Exception e) {
                            logger.error("Browser restart failed: " + e.getMessage());
                            throw new RuntimeException("Cannot continue without a browser", e);
                        }
                        sinceRestart = 0;
                        // after relaunch we're at the scheduler home; re-open this day
                        try { selectDay(date); waitForTableReload(); }
                        catch (Exception e) { recoverToScheduler(); }
                    }

                    boolean done = false;
                    for (int attempt = 0; attempt < ROW_FIND_ATTEMPTS && !done; attempt++) {
                        try {
                            dismissModalIfPresent();
                            selectDay(date);
                            waitForTableReload();

                            WebElement targetRow = findRowByKey(date, key);
                            if (targetRow == null) {
                                logger.warn("Row not yet found on " + date + " (attempt "
                                        + (attempt + 1) + "/" + ROW_FIND_ATTEMPTS + ")");
                                continue;
                            }

                            Appointment appt = buildAppointment(targetRow, date);
                            openAppointmentAndDownloadNote(targetRow, appt);

                            store.mark(key);
                            store.save();
                            audit.addRow(appt, FACILITY, appToday);
                            audit.flush();
                            downloaded++;
                            dayDownloaded++;
                            sinceRestart++;
                            done = true;
                            logger.info("Downloaded [" + downloaded + "] " + date
                                    + " | " + appt.patientName + " | MRN " + appt.mrn);
                        } catch (StaleElementReferenceException se) {
                            logger.warn("Stale on " + date + " (attempt " + (attempt + 1) + "); retrying");
                            recoverToScheduler();
                        } catch (Exception e) {
                            logger.error("Failed on " + date + " MRN " + mask(keyMrn(key))
                                    + " (" + e.getClass().getSimpleName() + "): " + e.getMessage());
                            recoverToScheduler();
                            break;
                        }
                    }
                    if (!done) {
                        errors++;
                        dayErrors++;
                        logger.error("UNRETRIEVED after retries on " + date
                                + " | MRN " + mask(keyMrn(key)));
                    }
                }

                if (dayDownloaded == expected) {
                    logger.info("Day " + date + " reconciled: expected " + expected
                            + ", downloaded " + dayDownloaded);
                } else {
                    logger.error("Day " + date + " INCOMPLETE: expected " + expected
                            + ", downloaded " + dayDownloaded + ", unretrieved " + dayErrors);
                }
            }
        } finally {
            store.pruneOlderThan(PRUNE_DAYS);
            store.save();
            audit.close();
            logger.info("Run complete. Downloaded: " + downloaded
                    + ", duplicates skipped: " + duplicates + ", errors: " + errors);
        }
    }

    /** Quit the current browser and launch a fresh, logged-in one. */
    private void restartBrowser() throws Exception {
        try { if (driver != null) driver.quit(); }
        catch (Exception e) { logger.warn("Quit during restart threw: " + e.getClass().getSimpleName()); }
        WebDriver fresh = loginAndOpen();
        setDriver(fresh);
        // confirm scheduler is up before continuing
        new WebDriverWait(driver, Duration.ofSeconds(30))
                .until(ExpectedConditions.presenceOfElementLocated(PICKER));
    }

    // =======================================================================
    //  HARVEST / ROW LOOKUP (stale-tolerant)
    // =======================================================================
    private List<Appointment> harvestWithRetry(LocalDate date) {
        for (int attempt = 0; attempt < 3; attempt++) {
            try {
                List<Appointment> out = new ArrayList<>();
                for (WebElement row : findCheckedRows()) out.add(buildAppointment(row, date));
                return out;
            } catch (StaleElementReferenceException se) {
                logger.warn("Stale during harvest on " + date + " (attempt " + (attempt + 1)
                        + "); re-selecting");
                try { selectDay(date); waitForTableReload(); } catch (Exception ignore) { }
            }
        }
        List<Appointment> out = new ArrayList<>();
        for (WebElement row : findCheckedRows()) out.add(buildAppointment(row, date));
        return out;
    }

    private WebElement findRowByKey(LocalDate date, String key) {
        for (WebElement row : findCheckedRows()) {
            try {
                if (rowKey(row, date).equals(key)) return row;
            } catch (StaleElementReferenceException se) { /* skip; later pass catches it */ }
        }
        return null;
    }

    // =======================================================================
    //  MODAL DISMISS + CLICK-WITH-RETRY
    // =======================================================================
    private boolean dismissModalIfPresent() {
        List<WebElement> bg = driver.findElements(MODAL_BG);
        boolean shown = !bg.isEmpty() && bg.get(0).isDisplayed();
        if (!shown) return false;

        String text = "";
        try {
            text = driver.findElement(By.cssSelector(".reveal-modal[style*='block'], .reveal-modal")).getText();
        } catch (Exception ignore) { }
        logger.warn("Modal appeared, clicking Exit: " + text.replaceAll("\\s+", " ").trim());

        try {
            List<WebElement> exit = driver.findElements(CONFIRM_EXIT);
            if (!exit.isEmpty() && exit.get(0).isDisplayed()) exit.get(0).click();
            else ((JavascriptExecutor) driver).executeScript(
                    "var b=document.getElementById('confirmYes'); if(b){b.click();}");
        } catch (Exception e) {
            logger.warn("Exit click failed: " + e.getClass().getSimpleName());
        }

        try {
            new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.invisibilityOfElementLocated(MODAL_BG));
        } catch (TimeoutException e) {
            logger.warn("Modal backdrop did not clear after Exit");
        }
        return true;
    }

    private void clickWithModalRetry(By locator) {
        try {
            wait.until(ExpectedConditions.elementToBeClickable(locator)).click();
        } catch (ElementClickInterceptedException e) {
            dismissModalIfPresent();
            wait.until(ExpectedConditions.elementToBeClickable(locator)).click();
        }
    }

    // =======================================================================
    //  PICKER HELPERS
    // =======================================================================
    private LocalDate readAppToday() {
        String text = wait.until(ExpectedConditions.visibilityOfElementLocated(TODAY_LINK)).getText();
        String datePart = text.substring(text.indexOf('-') + 1).trim();
        return LocalDate.parse(datePart, MDY);
    }

    private YearMonth readShownMonth() {
        String title = driver.findElement(CAL_TITLE).getText().trim();
        return YearMonth.parse(title, HDR);
    }

    private void navigateToMonth(YearMonth target) {
        for (int guard = 0; guard < 60; guard++) {
            YearMonth shown = readShownMonth();
            if (shown.equals(target)) return;
            String before = driver.findElement(CAL_TITLE).getText().trim();
            driver.findElement(shown.isAfter(target) ? PREV_ARROW : NEXT_ARROW).click();
            new WebDriverWait(driver, Duration.ofSeconds(5)).until(
                    d -> !d.findElement(CAL_TITLE).getText().trim().equals(before));
        }
        throw new IllegalStateException("Could not reach month " + target);
    }

    private void clickDay(int day) {
        By dayBtn = By.xpath(
                "//*[@id='scheduler-filter-bar-datepicker']" +
                "//button[not(contains(@class,'v-btn--fake-disabled'))]" +
                "[div[normalize-space(text())='" + day + "']]");
        wait.until(ExpectedConditions.visibilityOfElementLocated(PICKER));
        clickWithModalRetry(dayBtn);
    }

    private void selectDay(LocalDate date) {
        openCalendar();
        navigateToMonth(YearMonth.from(date));
        clickDay(date.getDayOfMonth());
    }

    // =======================================================================
    //  EXTRACTION
    // =======================================================================
    private List<WebElement> findCheckedRows() {
        return driver.findElements(By.xpath(
                "//button[contains(@class,'finalize-checkbox')]" +
                "[.//i[contains(@class,'fa-check-square-o')]]" +
                "/ancestor::tr[1]"));
    }

    private String rowKey(WebElement row, LocalDate calendarDate) {
        return cellText(row, "list-view-patient-number").trim() + ":" + calendarDate;
    }

    private static String keyMrn(String key) {
        int i = key.lastIndexOf(':');
        return i > 0 ? key.substring(0, i) : key;
    }

    private Appointment buildAppointment(WebElement row, LocalDate calendarDate) {
        String patientName = cellText(row, "event-title");
        String clinician   = cellText(row, "list-view-clinician");
        String visitType   = cellText(row, "list-view-visit-type");
        String apptDate    = cellText(row, "event-date");
        String dob         = cellText(row, "list-view-patient-dob");
        String mrn         = cellText(row, "list-view-patient-number").trim();

        String id = mrn + ":" + calendarDate;

        Appointment appt = new Appointment(id, calendarDate, patientName, clinician, visitType, apptDate);
        appt.mrn = mrn;
        appt.dob = dob;
        return appt;
    }

    private String cellText(WebElement row, String cssClass) {
        try {
            List<WebElement> cells = row.findElements(
                    By.cssSelector("td." + cssClass + ", span." + cssClass));
            return cells.isEmpty() ? "" : cells.get(0).getText().trim();
        } catch (StaleElementReferenceException e) {
            return "";
        }
    }

    private String pdfName(Appointment appt) {
        return sanitize(appt.patientName) + "_" + sanitize(appt.mrn) + "_" + appt.calendarDate + ".pdf";
    }

    // =======================================================================
    //  APPOINTMENT -> PROGRESS NOTE -> PDF
    // =======================================================================
    private Path openAppointmentAndDownloadNote(WebElement row, Appointment appt)
            throws IOException, InterruptedException {
        String listHandle = driver.getWindowHandle();
        Set<String> beforeAppt = driver.getWindowHandles();

        WebElement link = row.findElement(By.cssSelector("td.event-title a"));
        wait.until(ExpectedConditions.elementToBeClickable(link)).click();

        String apptTab = newHandleOrNull(beforeAppt);
        if (apptTab != null) driver.switchTo().window(apptTab);

        dismissModalIfPresent();

        Path pdf = downloadProgressNotePdf(appt);

        if (apptTab != null) {
            driver.close();
            driver.switchTo().window(listHandle);
        } else {
            driver.navigate().back();
            dismissModalIfPresent();
            wait.until(ExpectedConditions.presenceOfElementLocated(PICKER));
        }
        return pdf;
    }

    private Path downloadProgressNotePdf(Appointment appt) throws IOException, InterruptedException {
        String currentHandle = driver.getWindowHandle();
        Set<String> beforeDoc = driver.getWindowHandles();

        waitForOverlayToClear();
        dismissModalIfPresent();
        clickWithModalRetry(PROGRESS_NOTE_LINK);

        try {
            new WebDriverWait(driver, Duration.ofSeconds(5))
                    .until(ExpectedConditions.invisibilityOfElementLocated(MODAL_BG));
        } catch (TimeoutException ignore) { }

        waitForOverlayToClear();

        // Select the note whose date matches this record — NOT blindly the first item.
        By dateItem = dropdownItemForDate(appt.calendarDate);
        try {
            clickWithModalRetry(dateItem);
        } catch (TimeoutException e) {
            throw new IllegalStateException("No dropdown note matching "
                    + appt.calendarDate.format(DROPDOWN_DATE) + " for MRN " + mask(appt.mrn));
        }

        String docTab = newHandleOrNull(beforeDoc);
        if (docTab == null)
            throw new IllegalStateException("Progress-note document tab did not open");

        driver.switchTo().window(docTab);
        WebElement content = new WebDriverWait(driver, Duration.ofSeconds(15))
                .until(ExpectedConditions.presenceOfElementLocated(NOTE_CONTENT));

        String pageText = content.getText();
        Matcher m = NOTE_MRN.matcher(pageText);
        String pageMrn = m.find() ? m.group(1).trim() : "";
        if (!pageMrn.isEmpty() && !appt.mrn.isEmpty() && !pageMrn.equals(appt.mrn)) {
            driver.close();
            driver.switchTo().window(currentHandle);
            throw new IllegalStateException("MRN mismatch row=" + mask(appt.mrn)
                    + " note=" + mask(pageMrn) + " — not saving");
        }

        Path pdf = savePageAsPdf(appt);

        driver.close();
        driver.switchTo().window(currentHandle);
        return pdf;
    }

    private void waitForOverlayToClear() {
        try {
            new WebDriverWait(driver, Duration.ofSeconds(30))
                    .until(ExpectedConditions.invisibilityOfElementLocated(LOADING_OVERLAY));
        } catch (TimeoutException e) {
            logger.warn("Loading overlay did not clear within timeout");
        }
    }

    private Path savePageAsPdf(Appointment appt) throws IOException {
        PrintOptions po = new PrintOptions();
        po.setBackground(true);

        Pdf pdf = ((PrintsPage) driver).print(po);
        byte[] bytes = Base64.getDecoder().decode(pdf.getContent());

        if (bytes.length < 4 || bytes[0] != '%' || bytes[1] != 'P' || bytes[2] != 'D' || bytes[3] != 'F')
            throw new IOException("printToPDF did not produce a valid PDF");

        Path target = DOWNLOAD_DIR.resolve(pdfName(appt));
        Files.write(target, bytes);
        return target;
    }

    private String newHandleOrNull(Set<String> before) {
        try {
            return new WebDriverWait(driver, Duration.ofSeconds(5)).until(d -> {
                for (String h : d.getWindowHandles())
                    if (!before.contains(h)) return h;
                return null;
            });
        } catch (TimeoutException e) {
            return null;
        }
    }

    private void recoverToScheduler() {
        try {
            List<String> handles = new ArrayList<>(driver.getWindowHandles());
            for (int k = handles.size() - 1; k >= 1; k--) {
                driver.switchTo().window(handles.get(k));
                driver.close();
            }
            driver.switchTo().window(handles.get(0));
            dismissModalIfPresent();
            for (int hop = 0; hop < 3; hop++) {
                if (!driver.findElements(PICKER).isEmpty()
                        && driver.findElements(MODAL_BG).stream().noneMatch(WebElement::isDisplayed)) {
                    return;
                }
                driver.navigate().back();
                dismissModalIfPresent();
                try {
                    new WebDriverWait(driver, Duration.ofSeconds(5))
                            .until(ExpectedConditions.presenceOfElementLocated(PICKER));
                } catch (TimeoutException ignore) { }
            }
        } catch (Exception ignore) { }
    }

    private String sanitize(String s) {
        if (s == null) return "";
        return s.trim().replaceAll("[^a-zA-Z0-9._-]", "_");
    }

    private static String mask(String mrn) {
        if (mrn == null || mrn.length() <= 4) return "****";
        return "****" + mrn.substring(mrn.length() - 4);
    }

    // =======================================================================
    //  APP-SPECIFIC STUBS
    // =======================================================================
    private void openCalendar() { /* picker always rendered; no-op */ }

    private void waitForTableReload() {
        try {
            new WebDriverWait(driver, Duration.ofSeconds(20)).until(
                    ExpectedConditions.or(
                            ExpectedConditions.presenceOfElementLocated(By.cssSelector("td.list-view-status-check-in")),
                            ExpectedConditions.presenceOfElementLocated(By.cssSelector("td.event-title"))));
        } catch (TimeoutException ignored) { }
        try {
            int prev = -1;
            for (int i = 0; i < 12; i++) {
                int now = findCheckedRows().size();
                if (now == prev) return;
                prev = now;
                Thread.sleep(300);
            }
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }

    // =======================================================================
    //  ONE-TIME: backfill Visit Type into the corrected Excel from the front end.
    // =======================================================================
    public void backfillVisitTypes(Path correctedXlsx) throws IOException {
        try {
            new WebDriverWait(driver, Duration.ofSeconds(30))
                    .until(ExpectedConditions.presenceOfElementLocated(PICKER));
        } catch (TimeoutException e) {
            throw new IllegalStateException(
                "Scheduler date-picker not present — page not ready. URL: "
              + driver.getCurrentUrl(), e);
        }

        Workbook wb;
        try (InputStream in = Files.newInputStream(correctedXlsx)) {
            wb = WorkbookFactory.create(in);
        }
        Sheet sh = wb.getSheet("Audit") != null ? wb.getSheet("Audit") : wb.getSheetAt(0);

        Map<LocalDate, List<Row>> byDate = new LinkedHashMap<>();
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null) continue;
            String iso = cellStr(row.getCell(6));
            if (iso.isEmpty()) continue;
            byDate.computeIfAbsent(LocalDate.parse(iso), k -> new ArrayList<>()).add(row);
        }

        int filled = 0, missing = 0;
        for (Map.Entry<LocalDate, List<Row>> e : byDate.entrySet()) {
            LocalDate date = e.getKey();
            dismissModalIfPresent();
            selectDay(date);
            waitForTableReload();

            Map<String, String> visitByMrn = new HashMap<>();
            for (WebElement row : findCheckedRows()) {
                String mrn = cellText(row, "list-view-patient-number");
                String vt  = cellText(row, "list-view-visit-type");
                if (!mrn.isEmpty()) visitByMrn.putIfAbsent(mrn, vt);
            }

            for (Row xr : e.getValue()) {
                String mrn = cellStr(xr.getCell(3));
                String vt = visitByMrn.getOrDefault(mrn, "");
                Cell c = xr.getCell(2);
                if (c == null) c = xr.createCell(2);
                c.setCellValue(vt);
                if (vt.isEmpty()) { missing++; logger.warn("No visit type on " + date + " for a row"); }
                else filled++;
            }
        }

        try (OutputStream out = Files.newOutputStream(correctedXlsx)) { wb.write(out); }
        wb.close();
        logger.info("Visit-type backfill complete. Filled: " + filled + ", missing: " + missing);
    }

    private static String cellStr(Cell c) {
        if (c == null) return "";
        return c.getCellType() == CellType.STRING ? c.getStringCellValue().trim() : "";
    }

    // =======================================================================
    //  MODEL
    // =======================================================================
    static final class Appointment {
        final String id;
        final LocalDate calendarDate;
        final String patientName;
        final String clinician;
        final String visitType;
        final String apptDate;
        String mrn = "";
        String dob = "";

        Appointment(String id, LocalDate calendarDate, String patientName,
                    String clinician, String visitType, String apptDate) {
            this.id = id;
            this.calendarDate = calendarDate;
            this.patientName = patientName;
            this.clinician = clinician;
            this.visitType = visitType;
            this.apptDate = apptDate;
        }
    }

    // =======================================================================
    //  DEDUP STORE
    // =======================================================================
    static final class SeenStore {
        private final Path file;
        private final Map<String, String> seen;
        private static final Gson GSON = new Gson();

        SeenStore(Path file) throws IOException {
            this.file = file;
            if (Files.exists(file)) {
                Map<String, String> loaded = GSON.fromJson(
                        Files.readString(file),
                        new TypeToken<Map<String, String>>() {}.getType());
                this.seen = (loaded != null) ? loaded : new HashMap<>();
            } else {
                this.seen = new HashMap<>();
            }
        }

        boolean isNew(String id) { return !seen.containsKey(id); }
        void mark(String id)     { seen.put(id, LocalDate.now().toString()); }

        void pruneOlderThan(int days) {
            LocalDate cutoff = LocalDate.now().minusDays(days);
            seen.entrySet().removeIf(e -> LocalDate.parse(e.getValue()).isBefore(cutoff));
        }

        void save() throws IOException { Files.writeString(file, GSON.toJson(seen)); }
    }

    // =======================================================================
    //  EXCEL AUDIT
    // =======================================================================
    static final class ExcelTracker {
        private final Path file;
        private final Workbook wb;
        private final Sheet sheet;
        private static final String[] HEADERS = {
                "Patient Name", "Clinician Name", "Visit Type", "Patient MRN",
                "DOB", "Facility", "Date of Calendar Selected", "Date on Which it is traversed"
        };

        ExcelTracker(Path file) throws IOException {
            this.file = file;
            if (Files.exists(file)) {
                try (InputStream in = Files.newInputStream(file)) {
                    this.wb = WorkbookFactory.create(in);
                }
                this.sheet = wb.getSheet("Audit") != null ? wb.getSheet("Audit") : wb.createSheet("Audit");
                if (sheet.getRow(0) == null) writeHeader();
            } else {
                this.wb = new XSSFWorkbook();
                this.sheet = wb.createSheet("Audit");
                writeHeader();
            }
        }

        private void writeHeader() {
            Row header = sheet.createRow(0);
            CellStyle bold = wb.createCellStyle();
            Font f = wb.createFont();
            f.setBold(true);
            bold.setFont(f);
            for (int i = 0; i < HEADERS.length; i++) {
                Cell c = header.createCell(i);
                c.setCellValue(HEADERS[i]);
                c.setCellStyle(bold);
            }
        }

        void addRow(Appointment a, String facility, LocalDate traversedDate) {
            int rowNum = sheet.getLastRowNum() + 1;
            Row r = sheet.createRow(rowNum);
            r.createCell(0).setCellValue(a.patientName);
            r.createCell(1).setCellValue(a.clinician);
            r.createCell(2).setCellValue(a.visitType);
            r.createCell(3).setCellValue(a.mrn);
            r.createCell(4).setCellValue(a.dob);
            r.createCell(5).setCellValue(facility);
            r.createCell(6).setCellValue(a.calendarDate.toString());
            r.createCell(7).setCellValue(traversedDate.toString());
        }

        void flush() throws IOException {
            try (OutputStream out = Files.newOutputStream(file)) { wb.write(out); }
        }

        void close() throws IOException {
            for (int i = 0; i < HEADERS.length; i++) sheet.autoSizeColumn(i);
            flush();
            wb.close();
        }
    }

    // =======================================================================
    //  LOGIN
    // =======================================================================
    static WebDriver loginAndOpen() throws Exception {
        sel = new SeleniumUtils(projDirPath);
        WebDriver driver = sel.getDriver();

        utility = new Utility();
        String[] params = new String[]{"url", "username", "password"};
        HashMap<String, String> configs = utility.getConfig("config2.xml", params);
        String url = configs.get("url"),
               username = configs.get("username"),
               password = configs.get("password");

        driver.get(url);
        logger.info("Open url: " + url);
        sel.pauseClick(driver.findElement(By.id("Login")), 10);

        driver.findElement(By.id("Username")).sendKeys(username);
        driver.findElement(By.id("Password")).sendKeys(password);
        driver.findElement(By.id("Login")).click();
        Thread.sleep(4000);

        Select facility = new Select(driver.findElement(By.id("FacilityId")));
        facility.selectByVisibleText(FACILITY);
        logger.info("Selected facility: " + FACILITY);

        driver.findElement(By.id("Login")).click();
        Thread.sleep(8000);
        return driver;
    }

    // =======================================================================
    //  ENTRY POINT
    // =======================================================================

    @Test(priority = 1)
    public void runBot() throws Exception {
        WebDriver driver = null;
        CheckinextractorbotnewFinal bot = null;
        try {
            driver = loginAndOpen();
            bot = new CheckinextractorbotnewFinal(driver);
            bot.run();
        } finally {
            try { if (bot != null && bot.driver != null) bot.driver.quit(); }
            catch (Exception ignore) { }
        }
    }
    

    @AfterTest
	public static void fileCopy(ITestContext context) {


		int failed = context.getFailedTests().size();
    	if (failed > 0) {
        	System.out.println("Tests failed (" + failed + "). Skipping file copy.");
	        return;
    	}
	    String today = java.time.LocalDate.now()
	            .format(java.time.format.DateTimeFormatter.ofPattern("MM-dd-yyyy"));

	    // Filename-safe timestamp (no colons) e.g. 2026-06-25_12-00
	    String stamp = java.time.LocalDateTime.now()
	            .format(java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm"));

		Path sourceFile = AUDIT_FILE.toAbsolutePath();
		System.out.println("Source: " + sourceFile);

		Path destDir = Paths.get("\\\\10.172.192.34\\ai automation\\15 - Afra\\"+today+"\\");
		System.out.println("Destination dir: " + destDir);

	    try {
	        if (!Files.exists(destDir)) {
	            Files.createDirectories(destDir);
	        }

	        // Split original name into base + extension so the stamp goes before ".xlsx"
	        String originalName = sourceFile.getFileName().toString();
	        int dot = originalName.lastIndexOf('.');
	        String base = (dot == -1) ? originalName : originalName.substring(0, dot);
	        String ext  = (dot == -1) ? ""          : originalName.substring(dot); // includes the "."

	        String newName = base + " " + stamp + ext;

	        Path destFile = destDir.resolve(newName);
	        Files.copy(sourceFile, destFile, StandardCopyOption.REPLACE_EXISTING);
	        System.out.println("Excel file copied successfully to: " + destFile);
	    } catch (IOException e) {
	        e.printStackTrace();
	    }
	}
    
	@Test(priority = 2)
	public static void File_copy() {
		String today = java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("MM-dd-yyyy"));
		Path sourceDir = DOWNLOAD_DIR.toAbsolutePath();
		Path destDir = Paths.get("\\\\10.172.192.34\\ai automation\\15 - Afra\\"+today);
		try {
			if (!Files.exists(destDir)) {
				Files.createDirectories(destDir);
			}
			Files.walk(sourceDir).forEach(sourcePath -> {
				try {
					Path targetPath = destDir.resolve(sourceDir.relativize(sourcePath));
					if (Files.isDirectory(sourcePath)) {
						if (!Files.exists(targetPath)) {
							Files.createDirectories(targetPath);
						}
					} else {
						Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			});
			System.out.println("Directory copied successfully!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}


    
}