﻿///<Author>Nikolay Hsu</Author>
///<AuthorEmail>nikolay0514@gmail.com</AuthorEmail>
///<IDE>Visual Studio 2012</IDE>
///<Environment>Windows 7 Professional 64-bit</Environment>
///<Version>1.0</Version>
///<Description>
///This is program is designed to regularly send alert email to CanTeen Queenland members
///to notify them of which volunteers have licenses which are going to expire in a certain
///amount of time.
///The program reads volunteer data from a .csv file and examines each volunteer's license
///status, as well as his/her last activity date.
///The program then put together a alert list in the form of Excel spreadsheet(.xls) and sends
///email to the appointed email adress(es). Once finished or terminated, this .xls file will 
///be deleted.
///Upon execution, the program creates a log file(.txt) and stores it in the directory ./Log.
///If the program doesn't function as expected, the log can be useful for debuggin.
///Three items are required before running the program, namely qld_old_members.csv under directory
///C:\QLD Project\, as well as mail.html and MailList.txt under the program's directory.
///</Description>

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
    class Program {
        static void Main(string[] args) {
            AlertApp alertApp = new AlertApp();

            alertApp.run();
        }
    }
}
