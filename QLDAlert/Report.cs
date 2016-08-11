using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO;

namespace QLDAlert {
    class Report {
        private const string ALERT = "Alert";
        private const string ALL_VOLUNTEER = "All Volunteers";
        private const string INPUT_DATE_FORMAT = "M/d/yyyy";
        private const string OUTPUT_DATE_FORMAT = "dd/MM/yyyy";

        private Log log;
        private string reportPath;
        private Helper helper = new Helper();

        public Report(Log log, string reportPath) {
            this.log = log;
            this.reportPath = reportPath;
        }

        /// <summary>
        /// Create an xls file which contains of two sheets as the weekly report.
        /// Returns true if file has been created successfully.
        /// </summary>
        /// <returns>Returns boolean</returns>
        public bool initiate() {
            // deletes the existing report
            if (File.Exists(reportPath))
                File.Delete(reportPath);

            // creates a workbook
            XSSFWorkbook wb = new XSSFWorkbook();

            // creates a spreadsheet
            var sh = (XSSFSheet)wb.CreateSheet(ALERT);
            sh = (XSSFSheet)wb.CreateSheet(ALL_VOLUNTEER);

            try {
                using (var fs = new FileStream(reportPath, FileMode.Create, FileAccess.Write)) {
                    wb.Write(fs);
                }
            } catch (IOException ex) {
                log.appendLog("Failed to create report.\n" + ex.ToString(), -1);
                return false;
            }

            log.appendLog("Successfully created report.", 1);

            return true;
        }

        /// <summary>
        /// Examine each member. If the member has a license that will expire in certain peiord,
        /// put down his/her name and expiry date on the alert sheet.
        /// </summary>
        /// <param name="reportPath"></param>
        /// <param name="memberList"></param>
        public void writeReport(List<Member> memberList) {
            int lastRow = 0;
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet;
            DateTime noValue = DateTime.ParseExact("1/1/0001", INPUT_DATE_FORMAT, null);

            string[] headers = {"Name", "Blue Card Expiry Date", "Medical License Expiry Date",
                               "Medical Profession", "Police Check Expiry Date", "Last Activity",
                               "Contact ID"};

            try {
                FileStream fileStream =
                    new FileStream(reportPath, FileMode.Open, FileAccess.ReadWrite);
                workbook = new XSSFWorkbook(fileStream);
                fileStream.Close();
            } catch (Exception ex) {
                log.appendLog("Failed to open report.\n" + ex.ToString(), -1);
            }

            // Configure font and background color.
            XSSFCellStyle style = (XSSFCellStyle)workbook.CreateCellStyle();
            ICellStyle HeaderCellStyle = workbook.CreateCellStyle();
            XSSFFont font = (XSSFFont)workbook.CreateFont();
            font.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
            HeaderCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.DarkBlue.Index;
            HeaderCellStyle.FillPattern = FillPattern.SolidForeground;
            HeaderCellStyle.SetFont(font);
            HeaderCellStyle.Alignment = HorizontalAlignment.Center;

            sheet = (XSSFSheet)workbook.GetSheet(ALERT);
            sheet.CreateRow(lastRow);

            log.appendLog("Writing Alert sheet headers.", 1);

            // Write headers to the first row.
            for (int i = 0; i < headers.Length; i++) {
                sheet.GetRow(lastRow).CreateCell(i);
                sheet.GetRow(lastRow).GetCell(i).CellStyle = HeaderCellStyle;
                sheet.GetRow(lastRow).GetCell(i).SetCellValue(headers[i]);
            }

            lastRow++;

            log.appendLog("Checking volunteers' status.", 1);

            // Check each member's license expiry dates and date of last activity.
            foreach (Member mem in memberList) {

                bool isOnList = false;

                if (sheet.GetRow(lastRow) == null)
                    sheet.CreateRow(lastRow);

                if (!helper.isDateEmpty(mem.expDateBlueCard)) {
                    if (helper.isGoingToExpire(mem.expDateBlueCard)) {
                        sheet.GetRow(lastRow).CreateCell(1);
                        sheet.GetRow(lastRow).GetCell(1).SetCellValue(mem.expDateBlueCard.ToString(OUTPUT_DATE_FORMAT));
                        isOnList = true;
                    }
                }

                if (!helper.isDateEmpty(mem.expDateMedPro)) {
                    if (helper.isGoingToExpire(mem.expDateMedPro)) {
                        sheet.GetRow(lastRow).CreateCell(2);
                        sheet.GetRow(lastRow).CreateCell(3);
                        sheet.GetRow(lastRow).GetCell(2).SetCellValue(mem.expDateMedPro.ToString(OUTPUT_DATE_FORMAT));
                        sheet.GetRow(lastRow).GetCell(3).SetCellValue(mem.medProfession);
                        isOnList = true;
                    }
                }

                if (!helper.isDateEmpty(mem.expDatePoliceCheck)) {
                    if (helper.isGoingToExpire(mem.expDatePoliceCheck)) {
                        sheet.GetRow(lastRow).CreateCell(4);
                        sheet.GetRow(lastRow).GetCell(4).SetCellValue(mem.expDatePoliceCheck.ToString(OUTPUT_DATE_FORMAT));
                        isOnList = true;
                    }
                }

                if (!helper.isDateEmpty(mem.lastActivity)) {
                    if (!helper.isMemberActive(mem.lastActivity)) {
                        sheet.GetRow(lastRow).CreateCell(5);
                        sheet.GetRow(lastRow).GetCell(5).SetCellValue(mem.lastActivity.ToString(OUTPUT_DATE_FORMAT));
                        isOnList = true;
                    }
                }

                if (isOnList) {
                    sheet.GetRow(lastRow).CreateCell(0);
                    sheet.GetRow(lastRow).GetCell(0).SetCellValue(mem.name);
                    sheet.GetRow(lastRow).CreateCell(6);
                    sheet.GetRow(lastRow).GetCell(6).SetCellValue(mem.contactId);
                    lastRow++;
                }
            }

            for (int i = 0; i < headers.Length; i++) {
                sheet.AutoSizeColumn(i);
            }

            // Saves report.
            using (var fs = new FileStream(reportPath, FileMode.Create, FileAccess.Write)) {
                workbook.Write(fs);
            }

            log.appendLog("Finished writing alert report.", 1);
        }

        /// <summary>
        /// Write all volunteers to All Volunteers sheet.
        /// </summary>
        /// <param name="reportPath"></param>
        /// <param name="memberList"></param>
        /// <returns></returns>
        public void writeAllMember(List<Member> memberList) {
            int lastRow = 0;
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet;
            DateTime noValue = DateTime.ParseExact("1/1/0001", "M/d/yyyy", null);

            string[] headers = {"Name", "Email", "Postal Code", "Blue Card Expiry Date",
                                "Medical License Expiry Date", "Medical Profession", 
                                "Police Check Expiry Date", "Last Activity"};

            if (!File.Exists(reportPath)) {
                log.appendLog("Report doesn't exit.\n", -1);
                Environment.Exit(-1);
            }

            try {
                FileStream fileStream =
                    new FileStream(reportPath, FileMode.Open, FileAccess.ReadWrite);
                workbook = new XSSFWorkbook(fileStream);
                fileStream.Close();
            } catch (Exception ex) {
                log.appendLog("Failed to open report.\n" + ex.ToString(), 1);
            }

            XSSFCellStyle style = (XSSFCellStyle)workbook.CreateCellStyle();
            ICellStyle HeaderCellStyle = workbook.CreateCellStyle();
            HeaderCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.DarkBlue.Index;
            HeaderCellStyle.FillPattern = FillPattern.SolidForeground;
            XSSFFont font = (XSSFFont)workbook.CreateFont();
            font.Color = NPOI.HSSF.Util.HSSFColor.White.Index;
            HeaderCellStyle.SetFont(font);
            HeaderCellStyle.Alignment = HorizontalAlignment.Center;


            sheet = (XSSFSheet)workbook.GetSheet(ALL_VOLUNTEER);
            sheet.CreateRow(lastRow);

            log.appendLog("Writing All Volunteers sheet headers.", 1);

            for (int i = 0; i < headers.Length; i++) {
                if (sheet.GetRow(lastRow).GetCell(i) == null)
                    sheet.GetRow(lastRow).CreateCell(i);

                sheet.GetRow(lastRow).GetCell(i).CellStyle = HeaderCellStyle;

                sheet.GetRow(lastRow).GetCell(i).SetCellValue(headers[i]);
            }

            lastRow++;

            log.appendLog("Writing all volunteers to report.", 1);

            foreach (Member mem in memberList) {
                sheet.CreateRow(lastRow);

                for (int i = 0; i < headers.Length; i++) {
                    sheet.GetRow(lastRow).CreateCell(i);
                }

                sheet.GetRow(lastRow).GetCell(0).SetCellValue(mem.name);
                sheet.GetRow(lastRow).GetCell(1).SetCellValue(mem.email);
                sheet.GetRow(lastRow).GetCell(2).SetCellValue(mem.postcode);

                if (DateTime.Compare(noValue, mem.expDateBlueCard) != 0)
                    sheet.GetRow(lastRow).GetCell(3).SetCellValue(mem.expDateBlueCard.ToString(OUTPUT_DATE_FORMAT));

                if (DateTime.Compare(noValue, mem.expDateMedPro) != 0)
                    sheet.GetRow(lastRow).GetCell(4).SetCellValue(mem.expDateMedPro.ToString(OUTPUT_DATE_FORMAT));

                sheet.GetRow(lastRow).GetCell(5).SetCellValue(mem.medProfession);

                if (DateTime.Compare(noValue, mem.expDatePoliceCheck) != 0)
                    sheet.GetRow(lastRow).GetCell(6).SetCellValue(mem.expDatePoliceCheck.ToString(OUTPUT_DATE_FORMAT));

                if (DateTime.Compare(noValue, mem.lastActivity) != 0)
                    sheet.GetRow(lastRow).GetCell(7).SetCellValue(mem.lastActivity.ToString(OUTPUT_DATE_FORMAT));

                lastRow++;
            }

            for (int i = 0; i < headers.Length; i++) {
                sheet.AutoSizeColumn(i);
            }

            using (var fs = new FileStream(reportPath, FileMode.Create, FileAccess.Write)) {
                workbook.Write(fs);
            }

            log.appendLog("Finished writing all volunteers.", 1);
        }
    }
}
