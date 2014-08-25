using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    public class MailInfo
    {
        public Microsoft.Office.Interop.Outlook.MailItem Mail { get; set; }
        public string From { get; set; }
        public DateTime Date { get; set; }
        public string Subject { get; set; }
        public bool Skipped { get; set; }
    }

    public class MailBox
    {
        public Stack<MailInfo> DeleteList;
        public Stack<MailInfo> MailList;
        public Stack<MailInfo> NewList;
        public int NewCount;
        public List<string> Folders; 

        public MailBox()
        {
            DeleteList = new Stack<MailInfo>();
            MailList = new Stack<MailInfo>();
            NewList = new Stack<MailInfo>();
            Folders = new List<string>();
            NewCount = 0;
        }
    }
}
