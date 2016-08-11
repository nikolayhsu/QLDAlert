using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Web;

namespace QLDAlert {
    class AlertApp {

        #region Gobal and Const

        const string INPUT_DATE_FORMAT = "M/d/yyyy";
        const string OUTPUT_DATE_FORMAT = "dd/MM/yyyy";
        const string ALERT = "Alert";
        const string ALL_VOLUNTEER = "All Volunteers";
        const string CSV_PATH = @"R:\Vol\VOL_QLD.csv";
        const string MAIL_PATH = @".\mail.html";
        const string MAIL_RECIPIENT_LIST_PATH = @".\MailList.txt";

        private Log log;
        private Report report;
        private string reportPath;

        #endregion

        public void run() {
            string logDirectory = Directory.GetCurrentDirectory() + "\\Log";
            string reportDirectory = Directory.GetCurrentDirectory() + "\\Report";

            if (!Directory.Exists(logDirectory))
                Directory.CreateDirectory(logDirectory);

            if (!Directory.Exists(reportDirectory))
                Directory.CreateDirectory(reportDirectory);

            string logPath = @".\Log\Log_" + DateTime.Today.ToString("dd_MM_yyyy") + ".txt";

            log = new Log(logPath);

            if (!log.initiate())
                return;

            reportPath = @".\Report\Report_" + DateTime.Today.ToString("dd_MM_yyyy") + ".xlsx";

            report = new Report(log, reportPath);

            if (!report.initiate())
                return;

            List<Member> memberList = createList(CSV_PATH);

            if (memberList == null) {
                beforeExit(-1);
                return;
            }

            if (memberList.Count > 0) {
                report.writeReport(memberList);
                report.writeAllMember(memberList);
            } else {
                log.appendLog("No valid member data.", -1);
                beforeExit(-1);
                return;
            }

            Email mail = new Email(log);

            //mail.sendAlertMail(MAIL_PATH, reportPath);

            beforeExit(1);
        }

        /// <summary>
        /// Delete report and exit.
        /// </summary>
        private void beforeExit(int code) {
            try {
                log.appendLog("Deleting report file....", 1);

                if (File.Exists(reportPath))
                    File.Delete(reportPath);

                log.appendLog("Report file deleted.", 1);
            } catch (IOException ex) {
                log.appendLog("Failed to delete report file.", -1);
                log.appendLog(ex.ToString(), -1);
            }

            if(code == 1)
                log.appendLog("Mission accomplished. Normal exit.", 1);

            if(code == -1)
                log.appendLog("Abnormal exit.", -1);
        }

        #region Volunteer Data

        /// <summary>
        /// Load and read the csv file and store data in a list of Member.
        /// </summary>
        /// <param name="csvPath"></param>
        /// <returns></returns>
        private List<Member> createList(string csvPath) {
            int successCount = 0, failureCount = 0;

            List<Member> list = new List<Member>();

            log.appendLog("Loading CSV file.", 1);

            if (!File.Exists(csvPath)) {
                log.appendLog("CSV file doesn't exist.", -1);
                return null;
            }

            try {
                StreamReader reader = new StreamReader(File.OpenRead(csvPath));

                reader.ReadLine(); // Skip the headers

                while (!reader.EndOfStream) {
                    Member member = createMember(tokenize(reader.ReadLine()));

                    if (member == null) {
                        failureCount++;
                    } else {
                        successCount++;
                        list.Add(member);
                    }
                }

                reader.Close();
            } catch (Exception ex) {
                log.appendLog("Failed to open CSV file.\n" + ex.ToString(), -1);
                return null;
            }

            log.appendLog("Finished loading CSV file. " + (successCount + failureCount) + 
                " members in total. Failed to read " + failureCount + " member(s)." , 1);

            return list.OrderBy(o => o.name).ToList(); // Return sorted list
        }

        /// <summary>
        /// Tokenize a line read from the csv file and create a Member to store data.
        /// </summary>
        /// <param name="csvLine"></param>
        /// <returns></returns>
        private Member createMember(string[] values) {
            Member member = new Member();

            try {
                member.name = values[1] + " " + values[0];
                member.email = values[2];
                member.postcode = int.Parse(values[3]);

                if (values[4] != string.Empty)
                    member.expDateBlueCard = DateTime.ParseExact(values[4], INPUT_DATE_FORMAT, null);

                member.medProfession = values[5];

                if (values[6] != string.Empty)
                    member.expDateMedPro = DateTime.ParseExact(values[6], INPUT_DATE_FORMAT, null);

                if (values[7] != string.Empty)
                    member.expDatePoliceCheck = DateTime.ParseExact(values[7], INPUT_DATE_FORMAT, null);

                if (values[8] != string.Empty)
                    member.lastActivity = DateTime.ParseExact(values[8], INPUT_DATE_FORMAT, null);

                member.contactId = values[9];
            } catch (Exception ex) {
                log.appendLog("Failed to store data in Member.", -1);
                log.appendLog(ex.ToString(), -1);
                return null;
            }

            Console.WriteLine(member);

            return member;
        }

        private string[] tokenize(string csvLine) {
            string[] values = csvLine.Split(',');

            for (int i = 0; i < values.Length; i++) {
                values[i] = values[i].Replace("\"", string.Empty);
                values[i] = values[i].Trim();
            }

            return values;
        }

        #endregion
    }
}
