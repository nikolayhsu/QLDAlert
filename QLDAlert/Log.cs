using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace QLDAlert {
    class Log {
        private const string ERROR_HEADER = "[ERROR] ";
        private const string INFO_HEADER = "[INFO] ";

        private string logPath;

        public Log(string logPath) {
            this.logPath = logPath;
        }

        /// <summary>
        /// Creates a .txt log file. If log file exists, appends the time of execution.
        /// Returns false true if log file has been created successfully.
        /// </summary>
        /// <returns>Returns boolean</returns>
        public bool initiate() {
            try {
                if (!File.Exists(logPath)) {
                    using (StreamWriter sw = File.CreateText(logPath)) {
                        sw.WriteLine("Log " + DateTime.Today.ToString("d"));
                    }

                    appendLog("Successfully created log.", 1);
                } else {
                    string executionTime =
                        "===== " + DateTime.Now.ToString("HH:mm:ss") + " =====";

                    appendLog(executionTime, 0);
                }
            } catch (IOException ex) {
                Console.WriteLine("[ERROR] Failed to create log.\n" + ex.ToString());
                return false;
            }
            return true;
        }

        /// <summary>
        /// Prints message to screen and saves message to log.
        /// Returns true if message has been appended successfully.
        /// </summary>
        /// <param name="content"></param>
        /// <param name="msgNum"></param>
        /// <returns>Returns boolean</returns>
        public bool appendLog(string content, int msgNum) {
            string message = "";

            if (msgNum == 1)
                message = INFO_HEADER + content;
            else if (msgNum == -1)
                message = ERROR_HEADER + content;
            else
                message = content;

            Console.WriteLine(message);

            try {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(logPath, true)) {
                    file.WriteLine(message);
                }
            } catch (IOException ex) {
                Console.WriteLine("Failed to append log.\n" + ex.ToString());
                return false;
            }

            return true;
        }
    }
}
