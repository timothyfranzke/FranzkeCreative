using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;
using System.Text;

namespace OutlookAddIn1.Configuration
{
    public static class Recognizer
    {
        public static string SettingsGrammar = "settings";
        public static string[] Settings = {"settings"};
        public static string SettingsOptionsGrammar = "settings option";
        public static string[] SettingsOptions = {"mail", "voice"};
        public static string SettingsVoiceGrammar = "settings voice";
        public static string[] SettingsVoiceRate = {"talk faster", "talk slower", "faster", "slower"};
        public static string[] SettingsVoiceChange = {"change voice", "new voice"};
        public static string[] SettingsExit = {"exit"};
        public static string Yes = "yes";
        public static string No = "no";
        public static string GoGrammer = "go";
        public static string[] Continue = {"go", "continue"};
        public static string YesNoGrammar = "yesno";
        public static string[] YesNo = {"yes", "no"};
        public static string DeleteGrammar = "delete";
        public static string[] Delete = {"remove", "delete", "deeleet", "deeleat"};
        public static string MoveGrammar = "move";
        public static string[] Move = {"move", "move to folder", "folder", "folders", "options"};
        public static string SkipGrammar = "skip";
        public static string[] Skip = {"skip", "next", "necks"};
        public static string ReadGrammar = "read";
        public static string[] Read = {"read"};
        public static string RepeatGrammar = "repeat";
        public static string[] Repeat = {"repeat"};
        public static string ActionGrammar = "actions";
        public static string FolderGrammar = "folders";
        public static string[] FolderList = { "read folders", "what are my folder options", "folder options", "what folders" };

        public static string[] Folders(List<string> folders)
        {
            return folders.ToArray();
        }

        public static string[] Actions()
        {
            int i = 0;
            var actions = new string[Delete.Length + Move.Length + Skip.Length + Read.Length + Repeat.Length + Settings.Length + FolderList.Length];
            for (; i < Delete.Length; i++)
                actions[i] = Delete[i];
            for (int j = 0; j < Move.Length; i++, j++)
                actions[i] = Move[j];
            for (int k = 0; k < Skip.Length; i++, k++)
                actions[i] = Skip[k];
            for (int l = 0; l < Read.Length; i++, l++)
                actions[i] = Read[l];
            for (int m = 0; m < Repeat.Length; i++, m++)
                actions[i] = Repeat[m];
            for (int n = 0; n < Repeat.Length; i++, n++)
                actions[i] = Settings[n];
            for (int o = 0; o < FolderList.Length; i++, o++)
                actions[i] = FolderList[o];
            return actions;
        }

        public static string[] SettingsVoice()
        {
            int i = 0;
            var settings = new string[SettingsVoiceChange.Length + SettingsVoiceRate.Length + SettingsExit.Length];
            for (; i < SettingsVoiceChange.Length; i++)
                settings[i] = SettingsVoiceChange[i];
            for (int j = 0; j < SettingsVoiceRate.Length; i++, j++)
                settings[i] = SettingsVoiceRate[j];
            for (int k = 0; k < SettingsExit.Length; i++, k++)
                settings[i] = SettingsExit[k];
            return settings;
        }

        public static string[] SettingsOptionsList()
        {
            int i = 0;
            var settings = new string[SettingsOptions.Length + SettingsExit.Length];
            for (; i < SettingsOptions.Length; i++)
                settings[i] = SettingsOptions[i];
            for (int j = 0; j < SettingsExit.Length; i++, j++)
                settings[i] = SettingsExit[j];
            return settings;
        }
    }

    public static class GrammerNames
    {
        public static string Actions = "mailActions";

    }
}
