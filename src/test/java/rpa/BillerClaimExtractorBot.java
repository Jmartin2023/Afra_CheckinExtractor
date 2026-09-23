package rpa;

import objects.SeleniumUtils;
import objects.Utility;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.print.PrintOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.*;
import java.nio.file.*;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * ON-DEMAND BILLER CLAIM EXTRACTOR (separate from the 30-day reconciliation bot).
 *
 * Runs only when a biller drops an Excel file (path passed to main). For each row it
 * matches AFRA on Account Number + DOS, downloads the SOAP note PDF for a single match,
 * and annotates the biller's own workbook with an extraction status and the PDF filename.
 *
 * Matching:
 *   - Headers matched BY NAME (case-insensitive): "Account Number", "DOS".
 *   - Account Number: a leading "ZF" (case-insensitive) is stripped before matching; the
 *     portal shows the bare MRN. The biller's original cell is left untouched.
 *   - DOS cell may be a string or a numeric Excel date; both are handled -> LocalDate.
 *
 * Per-row status (added columns, biller rows preserved in place):
 *   EXTRACTED | NO_MATCH | MULTIPLE_MATCH | ERROR | BAD_DATE | BAD_ACCOUNT
 *
 * Reuses the reconciliation engine's behavior: login, day-select, find-row-by-MRN,
 * finalized-encounter modal (Exit), progress-note -> date-matched dropdown -> printToPDF,
 * browser restart every N downloads. NO dedup store (per current instruction).
 *
 * Output: annotated copy of the biller file + a timestamped PDF folder. Originals not
 * touched. PHI: outputs and logs hold PHI; protect them.
 *
 * Dependencies (Maven): selenium-java, poi-ooxml
 */
public class BillerClaimExtractorBot {
    static SeleniumUtils sel;
    static Utility utility;
    static String projDirPath;
    static Logger logger = LogManager.getLogger(BillerClaimExtractorBot.class);

    // ---- Config -----------------------------------------------------------
    private static final int    RESTART_EVERY = 70;
    private static final int    ROW_FIND_ATTEMPTS = 4;
    private static final String FACILITY = "Afra Wound Care Associates LLC";
    private static final Path   OUTPUT_ROOT = Paths.get("biller_runs");   // PHI: protect

    // Biller file to process. Overridable via arg[0]; falls back to this path.
    private static final Path   BILLER_FILE = Paths.get("C:\\Users\\jmartin\\eclipse-workspace\\NextGen Report Extraction\\not_in_db_claims.xlsx");

    // Header names in the biller file (matched case-insensitively).
    private static final String HDR_ACCOUNT = "Account Number";
    private static final String HDR_DOS     = "DOS";
    // Added annotation columns.
    private static final String HDR_STATUS  = "Extraction Status";
    private static final String HDR_PDF     = "PDF File";

    private WebDriver driver;
    private WebDriverWait wait;
    private Path pdfDir;   // per-run PDF output folder

    public BillerClaimExtractorBot(WebDriver driver) { setDriver(driver); }

    private void setDriver(WebDriver d) {
        this.driver = d;
        this.wait = new WebDriverWait(d, Duration.ofSeconds(20));
    }

    // ---- Locators ---------------------------------------------------------
    private static final By PICKER     = By.id("scheduler-filter-bar-datepicker");
    private static final By CAL_TITLE  = By.cssSelector(
            "#scheduler-filter-bar-datepicker .v-date-picker-header__value button");
    private static final By PREV_ARROW = By.cssSelector(
            "#scheduler-filter-bar-datepicker button[aria-label='Previous month']");
    private static final By NEXT_ARROW = By.cssSelector(
            "#scheduler-filter-bar-datepicker button[aria-label='Next month']");
    private static final By TODAY_LINK = By.id("scheduler-filter-today-button");
    private static final By LOADING_OVERLAY = By.cssSelector("div.overlay.overlay-front, div.overlay-fixed");
    private static final By PROGRESS_NOTE_LINK = By.cssSelector("a.progress-note-link");
    private static final By NOTE_CONTENT = By.id("divHideContent");
    private static final By MODAL_BG   = By.cssSelector("div.reveal-modal-bg");
    private static final By CONFIRM_EXIT = By.id("confirmYes");

    private static final Pattern NOTE_MRN = Pattern.compile("Patient Number:\\s*([A-Za-z0-9-]+)");

    private static final DateTimeFormatter HDR = DateTimeFormatter.ofPattern("MMMM yyyy", Locale.ENGLISH);
    private static final DateTimeFormatter DROPDOWN_DATE = DateTimeFormatter.ofPattern("M/d/yyyy", Locale.ENGLISH);

    private By dropdownItemForDate(LocalDate date) {
        String shown = date.format(DROPDOWN_DATE);
        return By.xpath(
            "//div[contains(@class,'v-list') and contains(@class,'v-list--dense')]" +
            "/div//a[contains(normalize-space(.),'" + shown + "')]");
    }

    // =======================================================================
    //  MAIN FLOW  — driven by the biller workbook
    // =======================================================================
    public void run(Path billerFile) throws IOException {
        if (!Files.exists(billerFile))
            throw new FileNotFoundException("Biller file not found: " + billerFile);

        // per-run output folder
        String stamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        Path runDir = OUTPUT_ROOT.resolve("run_" + stamp);
        pdfDir = runDir.resolve("pdfs");
        Files.createDirectories(pdfDir);

        Workbook wb;
        try (InputStream in = Files.newInputStream(billerFile)) {
            wb = WorkbookFactory.create(in);
        } catch (Exception e) {
            // "A biller file that cannot be read or does not follow the expected format."
            logger.error("Biller file cannot be read: " + e.getMessage());
            throw new IOException("Unreadable biller file", e);
        }

        Sheet sh = wb.getSheetAt(0);
        Row header = sh.getRow(sh.getFirstRowNum());
        if (header == null) throw new IOException("Biller file has no header row");

        int colAccount = findColumn(header, HDR_ACCOUNT);
        int colDos     = findColumn(header, HDR_DOS);
        if (colAccount < 0 || colDos < 0) {
            throw new IOException("Biller file missing required headers. Found account@"
                    + colAccount + ", dos@" + colDos
                    + " (need '" + HDR_ACCOUNT + "' and '" + HDR_DOS + "')");
        }
        int colStatus = ensureColumn(header, HDR_STATUS);
        int colPdf    = ensureColumn(header, HDR_PDF);

        int extracted = 0, noMatch = 0, multi = 0, errors = 0, badInput = 0, sinceRestart = 0;

        for (int r = sh.getFirstRowNum() + 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null) continue;

            String rawAccount = cellString(row.getCell(colAccount));
            String account = normalizeAccount(rawAccount);
            LocalDate dos = readDos(row.getCell(colDos));

            if (account.isEmpty()) { setStatus(row, colStatus, colPdf, "BAD_ACCOUNT", ""); badInput++; continue; }
            if (dos == null)       { setStatus(row, colStatus, colPdf, "BAD_DATE", "");    badInput++; continue; }

            // periodic browser restart
            if (sinceRestart >= RESTART_EVERY) {
                logger.info("Restarting browser after " + sinceRestart + " downloads");
                try { restartBrowser(); sinceRestart = 0; }
                catch (Exception e) { throw new RuntimeException("Cannot continue without a browser", e); }
            }

            String status, pdfFileName = "";
            try {
                dismissModalIfPresent();
                selectDay(dos);
                waitForTableReload();

                List<WebElement> matches = findRowsByMrn(dos, account);
                if (matches.isEmpty()) {
                    status = "NO_MATCH"; noMatch++;
                    logger.warn("NO_MATCH DOS " + dos + " acct " + mask(account));
                } else if (matches.size() > 1) {
                    // spec: report, do not resolve by assumption -> annotate, no download
                    status = "MULTIPLE_MATCH"; multi++;
                    logger.warn("MULTIPLE_MATCH (" + matches.size() + ") DOS " + dos + " acct " + mask(account));
                } else {
                    Appointment appt = buildAppointment(matches.get(0), dos);
                    Path pdf = openAppointmentAndDownloadNote(matches.get(0), appt);
                    pdfFileName = pdf.getFileName().toString();
                    status = "EXTRACTED"; extracted++; sinceRestart++;
                    logger.info("EXTRACTED DOS " + dos + " | " + appt.patientName
                            + " | MRN " + appt.mrn + " -> " + pdfFileName);
                }
            } catch (Exception e) {
                status = "ERROR"; errors++;
                logger.error("ERROR DOS " + dos + " acct " + mask(account)
                        + " (" + e.getClass().getSimpleName() + "): " + e.getMessage());
                recoverToScheduler();
            }
            setStatus(row, colStatus, colPdf, status, pdfFileName);
        }

        // write annotated copy (do not overwrite the biller original)
        Path annotated = runDir.resolve("annotated_" + billerFile.getFileName());
        try (OutputStream out = Files.newOutputStream(annotated)) { wb.write(out); }
        wb.close();

        logger.info("Biller run complete. extracted=" + extracted + " noMatch=" + noMatch
                + " multiMatch=" + multi + " errors=" + errors + " badInput=" + badInput);
        logger.info("Annotated file -> " + annotated);
        logger.info("PDFs -> " + pdfDir);
    }

    // =======================================================================
    //  BILLER-FILE HELPERS
    // =======================================================================
    private int findColumn(Row header, String name) {
        for (int c = header.getFirstCellNum(); c < header.getLastCellNum(); c++) {
            String h = cellString(header.getCell(c));
            if (h.equalsIgnoreCase(name.trim())) return c;
        }
        return -1;
    }

    /** Find a column by name, or append it at the end and return its index. */
    private int ensureColumn(Row header, String name) {
        int existing = findColumn(header, name);
        if (existing >= 0) return existing;
        int idx = Math.max(0, (int) header.getLastCellNum());  // append
        header.createCell(idx).setCellValue(name);
        return idx;
    }

    private void setStatus(Row row, int colStatus, int colPdf, String status, String pdf) {
        cell(row, colStatus).setCellValue(status);
        cell(row, colPdf).setCellValue(pdf);
    }

    private Cell cell(Row row, int idx) {
        Cell c = row.getCell(idx);
        return (c != null) ? c : row.createCell(idx);
    }

    /** Strip a leading case-insensitive "ZF" (with optional surrounding space). */
    private String normalizeAccount(String raw) {
        if (raw == null) return "";
        String s = raw.trim();
        if (s.length() >= 2 && s.substring(0, 2).equalsIgnoreCase("ZF")) s = s.substring(2);
        return s.trim();
    }

    /** DOS cell may be text or a numeric Excel date. Returns null if unparseable. */
    private LocalDate readDos(Cell c) {
        if (c == null) return null;
        try {
            if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
                return c.getLocalDateTimeCellValue().toLocalDate();
            }
            String s = cellString(c).trim();
            if (s.isEmpty()) return null;
            // try common textual formats
            for (DateTimeFormatter fmt : new DateTimeFormatter[]{
                    DateTimeFormatter.ISO_LOCAL_DATE,                        // 2026-06-29
                    DateTimeFormatter.ofPattern("M/d/yyyy", Locale.ENGLISH), // 6/29/2026
                    DateTimeFormatter.ofPattern("MM/dd/yyyy", Locale.ENGLISH)}) {
                try { return LocalDate.parse(s, fmt); } catch (Exception ignore) { }
            }
            return null;
        } catch (Exception e) {
            return null;
        }
    }

    private String cellString(Cell c) {
        if (c == null) return "";
        switch (c.getCellType()) {
            case STRING:  return c.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c))
                    return c.getLocalDateTimeCellValue().toLocalDate().toString();
                // avoid scientific notation / trailing .0 on account-like numbers
                double d = c.getNumericCellValue();
                if (d == Math.floor(d)) return String.valueOf((long) d);
                return String.valueOf(d);
            case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
            case FORMULA:
                try { return c.getStringCellValue().trim(); } catch (Exception e) { return ""; }
            default:      return "";
        }
    }

    // =======================================================================
    //  MATCHING  — rows on `date` whose MRN equals `account`
    // =======================================================================
    private List<WebElement> findRowsByMrn(LocalDate date, String account) {
        List<WebElement> out = new ArrayList<>();
        for (WebElement row : findCheckedRows()) {
            try {
                String mrn = cellText(row, "list-view-patient-number").trim();
                if (mrn.equalsIgnoreCase(account)) out.add(row);
            } catch (StaleElementReferenceException ignore) { }
        }
        return out;
    }

    // Note: reconciliation used only CHECKED rows. Biller matching should consider ALL
    // rows on the day (a claim may exist regardless of check-in state). If the portal
    // only exposes MRN on checked rows, revert to findCheckedRows-only. VERIFY.
    private List<WebElement> findCheckedRows() {
        // All appointment rows on the day (broader than the reconciliation bot).
        List<WebElement> all = driver.findElements(By.cssSelector("tr"));
        List<WebElement> rows = new ArrayList<>();
        for (WebElement r : all) {
            if (!r.findElements(By.cssSelector("td.list-view-patient-number")).isEmpty()) rows.add(r);
        }
        return rows;
    }

    // =======================================================================
    //  BROWSER RESTART
    // =======================================================================
    private void restartBrowser() throws Exception {
        try { if (driver != null) driver.quit(); } catch (Exception ignore) { }
        setDriver(loginAndOpen());
        new WebDriverWait(driver, Duration.ofSeconds(30))
                .until(ExpectedConditions.presenceOfElementLocated(PICKER));
    }

    // =======================================================================
    //  MODAL / CLICK
    // =======================================================================
    private boolean dismissModalIfPresent() {
        List<WebElement> bg = driver.findElements(MODAL_BG);
        if (bg.isEmpty() || !bg.get(0).isDisplayed()) return false;
        String text = "";
        try { text = driver.findElement(By.cssSelector(".reveal-modal[style*='block'], .reveal-modal")).getText(); }
        catch (Exception ignore) { }
        logger.warn("Modal appeared, clicking Exit: " + text.replaceAll("\\s+", " ").trim());
        try {
            List<WebElement> exit = driver.findElements(CONFIRM_EXIT);
            if (!exit.isEmpty() && exit.get(0).isDisplayed()) exit.get(0).click();
            else ((JavascriptExecutor) driver).executeScript(
                    "var b=document.getElementById('confirmYes'); if(b){b.click();}");
        } catch (Exception e) { logger.warn("Exit click failed: " + e.getClass().getSimpleName()); }
        try {
            new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.invisibilityOfElementLocated(MODAL_BG));
        } catch (TimeoutException e) { logger.warn("Modal backdrop did not clear after Exit"); }
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
    //  PICKER
    // =======================================================================
    private LocalDate readAppToday() {
        String text = wait.until(ExpectedConditions.visibilityOfElementLocated(TODAY_LINK)).getText();
        return LocalDate.parse(text.substring(text.indexOf('-') + 1).trim(),
                DateTimeFormatter.ofPattern("MM/dd/yyyy", Locale.ENGLISH));
    }

    private YearMonth readShownMonth() {
        return YearMonth.parse(driver.findElement(CAL_TITLE).getText().trim(), HDR);
    }

    private void navigateToMonth(YearMonth target) {
        for (int guard = 0; guard < 120; guard++) {   // biller DOS can be far in the past
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
        By dayBtn = By.xpath("//*[@id='scheduler-filter-bar-datepicker']" +
                "//button[not(contains(@class,'v-btn--fake-disabled'))]" +
                "[div[normalize-space(text())='" + day + "']]");
        wait.until(ExpectedConditions.visibilityOfElementLocated(PICKER));
        clickWithModalRetry(dayBtn);
    }

    private void selectDay(LocalDate date) {
        navigateToMonth(YearMonth.from(date));
        clickDay(date.getDayOfMonth());
    }

    // =======================================================================
    //  APPOINTMENT -> SOAP NOTE PDF
    // =======================================================================
    private Appointment buildAppointment(WebElement row, LocalDate date) {
        String patientName = cellText(row, "event-title");
        String clinician   = cellText(row, "list-view-clinician");
        String visitType   = cellText(row, "list-view-visit-type");
        String apptDate    = cellText(row, "event-date");
        String dob         = cellText(row, "list-view-patient-dob");
        String mrn         = cellText(row, "list-view-patient-number").trim();
        Appointment a = new Appointment(mrn + ":" + date, date, patientName, clinician, visitType, apptDate);
        a.mrn = mrn; a.dob = dob;
        return a;
    }

    private Path openAppointmentAndDownloadNote(WebElement row, Appointment appt)
            throws IOException, InterruptedException {
        String listHandle = driver.getWindowHandle();
        Set<String> beforeAppt = driver.getWindowHandles();

        WebElement link = row.findElement(By.cssSelector("td.event-title a"));
        wait.until(ExpectedConditions.elementToBeClickable(link)).click();

        String apptTab = newHandleOrNull(beforeAppt);
        if (apptTab != null) driver.switchTo().window(apptTab);
        dismissModalIfPresent();

        Path pdf = downloadSoapNotePdf(appt);

        if (apptTab != null) { driver.close(); driver.switchTo().window(listHandle); }
        else { driver.navigate().back(); dismissModalIfPresent();
               wait.until(ExpectedConditions.presenceOfElementLocated(PICKER)); }
        return pdf;
    }

    private Path downloadSoapNotePdf(Appointment appt) throws IOException, InterruptedException {
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

        By dateItem = dropdownItemForDate(appt.calendarDate);
        try { clickWithModalRetry(dateItem); }
        catch (TimeoutException e) {
            throw new IllegalStateException("No dropdown note matching "
                    + appt.calendarDate.format(DROPDOWN_DATE) + " for MRN " + mask(appt.mrn));
        }

        String docTab = newHandleOrNull(beforeDoc);
        if (docTab == null) throw new IllegalStateException("SOAP note document tab did not open");
        driver.switchTo().window(docTab);
        WebElement content = new WebDriverWait(driver, Duration.ofSeconds(15))
                .until(ExpectedConditions.presenceOfElementLocated(NOTE_CONTENT));

        Matcher m = NOTE_MRN.matcher(content.getText());
        String pageMrn = m.find() ? m.group(1).trim() : "";
        if (!pageMrn.isEmpty() && !appt.mrn.isEmpty() && !pageMrn.equals(appt.mrn)) {
            driver.close(); driver.switchTo().window(currentHandle);
            throw new IllegalStateException("MRN mismatch row=" + mask(appt.mrn)
                    + " note=" + mask(pageMrn) + " — not saving");
        }

        Path pdf = savePageAsPdf(appt);
        driver.close(); driver.switchTo().window(currentHandle);
        return pdf;
    }

    private void waitForOverlayToClear() {
        try {
            new WebDriverWait(driver, Duration.ofSeconds(30))
                    .until(ExpectedConditions.invisibilityOfElementLocated(LOADING_OVERLAY));
        } catch (TimeoutException e) { logger.warn("Loading overlay did not clear within timeout"); }
    }

    private Path savePageAsPdf(Appointment appt) throws IOException {
        PrintOptions po = new PrintOptions();
        po.setBackground(true);
        Pdf pdf = ((PrintsPage) driver).print(po);
        byte[] bytes = Base64.getDecoder().decode(pdf.getContent());
        if (bytes.length < 4 || bytes[0] != '%' || bytes[1] != 'P' || bytes[2] != 'D' || bytes[3] != 'F')
            throw new IOException("printToPDF did not produce a valid PDF");
        String fileName = sanitize(appt.patientName) + "_" + sanitize(appt.mrn) + "_" + appt.calendarDate + ".pdf";
        Path target = pdfDir.resolve(fileName);
        Files.write(target, bytes);
        return target;
    }

    private String newHandleOrNull(Set<String> before) {
        try {
            return new WebDriverWait(driver, Duration.ofSeconds(5)).until(d -> {
                for (String h : d.getWindowHandles()) if (!before.contains(h)) return h;
                return null;
            });
        } catch (TimeoutException e) { return null; }
    }

    private void recoverToScheduler() {
        try {
            List<String> handles = new ArrayList<>(driver.getWindowHandles());
            for (int k = handles.size() - 1; k >= 1; k--) {
                driver.switchTo().window(handles.get(k)); driver.close();
            }
            driver.switchTo().window(handles.get(0));
            dismissModalIfPresent();
            for (int hop = 0; hop < 3; hop++) {
                if (!driver.findElements(PICKER).isEmpty()
                        && driver.findElements(MODAL_BG).stream().noneMatch(WebElement::isDisplayed)) return;
                driver.navigate().back();
                dismissModalIfPresent();
                try {
                    new WebDriverWait(driver, Duration.ofSeconds(5))
                            .until(ExpectedConditions.presenceOfElementLocated(PICKER));
                } catch (TimeoutException ignore) { }
            }
        } catch (Exception ignore) { }
    }

    private String cellText(WebElement row, String cssClass) {
        try {
            List<WebElement> cells = row.findElements(By.cssSelector("td." + cssClass + ", span." + cssClass));
            return cells.isEmpty() ? "" : cells.get(0).getText().trim();
        } catch (StaleElementReferenceException e) { return ""; }
    }

    private void waitForTableReload() {
        try {
            new WebDriverWait(driver, Duration.ofSeconds(20)).until(
                    ExpectedConditions.or(
                            ExpectedConditions.presenceOfElementLocated(By.cssSelector("td.list-view-patient-number")),
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
        } catch (InterruptedException e) { Thread.currentThread().interrupt(); }
    }

    private String sanitize(String s) { return (s == null) ? "" : s.trim().replaceAll("[^a-zA-Z0-9._-]", "_"); }
    private static String mask(String mrn) {
        if (mrn == null || mrn.length() <= 4) return "****";
        return "****" + mrn.substring(mrn.length() - 4);
    }

    // =======================================================================
    //  MODEL
    // =======================================================================
    static final class Appointment {
        final String id; final LocalDate calendarDate;
        final String patientName, clinician, visitType, apptDate;
        String mrn = "", dob = "";
        Appointment(String id, LocalDate calendarDate, String patientName,
                    String clinician, String visitType, String apptDate) {
            this.id = id; this.calendarDate = calendarDate; this.patientName = patientName;
            this.clinician = clinician; this.visitType = visitType; this.apptDate = apptDate;
        }
    }

    // =======================================================================
    //  LOGIN
    // =======================================================================
    static WebDriver loginAndOpen() throws Exception {
        sel = new SeleniumUtils(projDirPath);
        WebDriver driver = sel.getDriver();
        utility = new Utility();
        HashMap<String, String> configs = utility.getConfig("config2.xml",
                new String[]{"url", "username", "password"});
        driver.get(configs.get("url"));
        logger.info("Open url: " + configs.get("url"));
        sel.pauseClick(driver.findElement(By.id("Login")), 10);
        driver.findElement(By.id("Username")).sendKeys(configs.get("username"));
        driver.findElement(By.id("Password")).sendKeys(configs.get("password"));
        driver.findElement(By.id("Login")).click();
        Thread.sleep(4000);
        new Select(driver.findElement(By.id("FacilityId"))).selectByVisibleText(FACILITY);
        logger.info("Selected facility: " + FACILITY);
        driver.findElement(By.id("Login")).click();
        Thread.sleep(8000);
        return driver;
    }

    // =======================================================================
    //  ENTRY POINT  — pass the biller file path as arg[0]
    // =======================================================================
    public static void main(String[] args) throws Exception {
        // Use arg[0] if supplied, otherwise the hardcoded BILLER_FILE.
        Path billerFile = (args.length >= 1) ? Paths.get(args[0]) : BILLER_FILE;
        logger.info("Biller file: " + billerFile);
        WebDriver driver = null;
        BillerClaimExtractorBot bot = null;
        try {
            driver = loginAndOpen();
            bot = new BillerClaimExtractorBot(driver);
            bot.run(billerFile);
        } finally {
            try { if (bot != null && bot.driver != null) bot.driver.quit(); }
            catch (Exception ignore) { }
        }
    }
}