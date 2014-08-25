using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OutlookAddIn1.Configuration
{
    public static class RuntimeConfig
    {
        public static bool EmptyBox = false;
        public static bool Start = false;
        public static bool First = false;
        public static bool Repeat = false;
        public static bool MailNew = false;
        public static bool Move = false;
        public static bool Delete = false;
        public static bool Yes = false;
        public static bool No = false;
        public static bool Processed = false;
        public static bool Go = false;
        public static bool Skip = false;
        public static bool Read = false;
        public static bool Settings = false;
        public static bool ReadFolders = false;
        public static string FolderName;
        public static int SkippedCount = 0;
        public static int GrammarIndex = 0;
        
        #region readerActions

        public static bool ReaderActionsFolderNames = false;
        #endregion

        #region actions

        public static bool ActionsRepeat = false;
        public static bool ActionsMove = false;
        public static bool ActionsMoveFolders = false;
        public static bool ActionsSkip = false;
        public static bool ActionsDelete = false;
        public static bool ActionsRead = false;


        #endregion

        #region settings

        public static bool SettingsMail = false;
        public static bool SettingsMailNew = false;
        public static bool SettingsMailInbox = false;
        public static bool SettingsVoice = false;
        public static bool SettingsVoiceChange = false;
        public static bool SettingsVoiceRate = false;
        public static bool SettingsVoiceRateIncrease = false;
        public static bool SettingsVoiceRateDecrease = false;
        public static bool SettingsExit = false;

        #endregion

        #region

        public static bool SpeakerStop = false;       

        #endregion

        #region mailbox

        public static bool MailboxEmpty = false;
        public static bool MailboxFirst = false;

        #endregion


        public static void Reset()
        {
            FolderName = String.Empty;
            Skip = false;
            Go = false;
            EmptyBox = false;
            Start = false;
            First = false;
            Repeat = false;
            MailNew = false;
            Delete = false;
            Yes = false;
            No = false;
            Processed = false;

            ActionsDelete = false;
            ActionsMove = false;
            ActionsMoveFolders = false;
            ActionsRead = false;
            ActionsRepeat = false;
            ActionsSkip = false;

            Settings = false;
            SettingsExit = false;
            SettingsMail = false;
            SettingsMailInbox = false;
            SettingsMailNew = false;
            SettingsVoice = false;
            SettingsVoiceChange = false;
            SettingsVoiceRate = false;
            SettingsVoiceRateDecrease = false;
            SettingsVoiceRateIncrease = false;

            MailboxEmpty = false;
            MailboxFirst = false;
        }
    }
}
