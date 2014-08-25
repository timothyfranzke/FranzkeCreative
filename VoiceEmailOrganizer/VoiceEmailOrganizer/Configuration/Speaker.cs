using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace VoiceEmailOrganizer.Configuration
{
    public static class Questions
    {
        public static string Start = "would you like me to start";
        public static string Folder = "which folder";
        public static string Proceed = "how would you like to proceed";
        public static string NewMail = "Should I read your new message";
        public static string Continue = "should I continue ";
    }

    public static class Confirmation
    {
        public static string Action(string action)
        {
            return string.Format("are you sure you want to {0} this message", action);
        }
    }

    public static class Statements
    {

        public static string Paused = "The system will be paused";
        public static string Go = "Say go to resoom";
        public static string ReadNew = "Should I read your new message";
        public static string FirstMessage = "";
        public static string RepeatMessage = "this message was is ";
        public static string NextMessage = "";
        public static string DeleteMessage = "message deleted ";
        public static string NoUnreadMessages = "";
        public static string FolderNames = "your folders are";

        public static string MessageFirst = "first unread message";

        #region Actions

        public static string ActionsRepeat = "this message is";
        public static string ActionsDelete = "message deleted";
        public static string ActionsMoveQuestion = "which folder";
        public static string ActionMoveFolderList(List<string> folders)
        {
            var folderList = new StringBuilder();
            folderList.Append("your folder choices are ");
            foreach (var folder in folders)
            {
                folderList.Append(folder.ToString());
            }
            return folderList.ToString();
        }

        public static string ActionMoveConfirmation(string folder)
        {
            return string.Format("this message has been moved to the {0} folder", folder);
        }
        public static string ActionsSkip = "this message will remain unread";
        public static string Actions(string action)
        {
            return string.Format("this message will be {0}", action);
        }

        #endregion

        #region Settings

        public static string Settings = "settings. say voice or mail";
        public static string SettingsMail = "mail settings. new or inbox";
        public static string SettingsMailNew = "do you want me to read new messages";
        public static string SettingsMailInbox = "do you want me to read only unread messages";
        public static string SettingsVoice = "say talk faster talk slower or change voice";
        public static string SettingsVoiceChange(string name)
        {
            return string.Format("hello my name is {0}, do you like my voice", name);
        }
        public static string SettingVoiceChangeConfirmation = "I am now your voice";
        public static string SettingsVoiceSpeedConfirmation(string rate)
        {
            return string.Format("my speech rate has been {0}", rate);
        }
        public static string SettingsVoiceSpeedIncreased = "increased";
        public static string SettingsVoiceSpeedDecreased = "decreased";
        public static string SettingsExit = "say exit to read the next mail item";
        public static string SettingsConfirmation = "";

        #endregion

        public static string Action(string action)
        {
            return string.Format("this e mail will be {0}", action);
        }

        #region MailBox

        public static string MailBox;
        public static string MailBoxUnreadFirst = "first unread message";
        public static string MailBoxUnreadNext = "next unread message";
        public static string MailBoxUnreadEmpty = "there are no unread messages in your mailbox";
        public static string MailBoxNewMail(int messageNumbers)
        {
            return string.Format("You have received {0} new {1}", messageNumbers,
                messageNumbers > 1 ? "messages" : "message");
        }

        public static string MailBoxNewMailQuestion = "should I read your new mail";
        public static string MailBoxMessageFrom(string sender)
        {
            return string.Format("from {0}", sender);
        }
        public static string MailBoxMessageSubject(string subject)
        {
            return string.Format("with a subject of {0}", subject);
        }

        public static string MailBoxMessageDate(string date)
        {
            return string.Format("received {0}", date);
        }

        #endregion

        #region speaker

        public static string SpeakerQuestionStart = "would you like me to start";
        public static string SpeakerPause = "The system will be paused";
        public static string SpeakerGo = "say go to rezoom";
        public static string SpeakerNoVoiceRecognition = "you do not have microsoft voice recognition on this computer.  This add in cannot operate without it";

        #endregion

    }

    public static class SingleStatements
    {
        public static string Delete = "delete";
        public static string Deleted = "deleted";
        public static string Skipped = "skipped";
    }
}
