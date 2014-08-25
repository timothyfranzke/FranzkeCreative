using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Windows.Forms;
using System.Xml.Linq;
using OutlookAddIn1.Configuration;
using OutlookAddIn1.Repository;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Timer = System.Timers.Timer;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private SpeechSynthesizer synth = new SpeechSynthesizer();
        private SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();
        public Outlook.MAPIFolder inBox;
        public Outlook.Items items;
        Timer timer = new Timer(AppConfig.TimerLength);
        public MailInfo MailInfo;
        public Outlook.MailItem OutlookMailItem = null;
        public MailBox MailBox;
        public GrammarRepository GrammarRepo;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var installedRecognizers = SpeechRecognitionEngine.InstalledRecognizers();
            foreach (var rec in installedRecognizers)
            {
                var reco = rec;
                if (rec.Name == AppConfig.MicrosoftVoiceRecognizerName)
                    AppConfig.SpeakerNoVoiceRecognition = true;
            }
            if (AppConfig.SpeakerNoVoiceRecognition)
            {
                recognizer.SetInputToDefaultAudioDevice();
                Application.NewMailEx += Application_NewMailEx;
                recognizer.SpeechRecognized += recognizer_SpeechRecognized;
               
                inBox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.Folders fold = inBox.Folders;
                GrammarRepo = new GrammarRepository();
                MailBox = new MailBox();
                synth.Rate = AppConfig.SpeakerSpeechRate;
                items = (Outlook.Items)inBox.Items.Restrict("[unread]=true");

                var accounts = Application.Session.Accounts;

                FillMailList();
                CreateFoldersGrammar();
                LoadGrammars();
                DisableGrammars();
                if (MailBox.MailList.Count == 0)
                {
                    RuntimeConfig.EmptyBox = true;
                    synth.Speak(Statements.NoUnreadMessages);
                }
                else
                {
                    RuntimeConfig.Start = true;
                    synth.Speak(Questions.Start);
                    EnableRecognizer(GrammarRepo.YesNo);
                }
            }
            else
            {
                synth.Speak(Statements.SpeakerNoVoiceRecognition);
            }
        }

        private void FillMailList()
        {
            items = (Outlook.Items)inBox.Items.Restrict("[unread]=true");
            items.GetFirst();
            var holder = items.GetFirst();
            var mItem = new MailInfo() { Mail = holder };
            MailBox.MailList.Push(mItem);

            for (int i = 0; i < items.Count - 1; i++)
            {
                Outlook.MailItem mailItem = items.GetNext();
                mItem = new MailInfo() { Mail = mailItem };
                MailBox.MailList.Push(mItem);
            }
            MailBox.Folders = GetFolders();
        }

        private void DisableGrammars()
        {
            for (int i = 0; i < recognizer.Grammars.Count; i++)
            {
                recognizer.Grammars[i].Enabled = false;
            }
        }

        private List<string> GetFolders()
        {
            var folders = new List<string>();
            foreach (Outlook.MAPIFolder item in inBox.Folders)
            {
                folders.Add(item.Name.ToString());
            }
            return folders;
        }

        private void CreateFoldersGrammar()
        {
            GrammarRepo.Folders = GrammarRepo.CreateGrammar(MailBox.Folders.ToArray(), "folders");
        }

        private void VoiceChoice()
        {
            var voices = synth.GetInstalledVoices();
            if (AppConfig.SpeakervoiceIndexNew + 1 == voices.Count)
            {
                AppConfig.SpeakervoiceIndexNew = 0;
            }
            else
            {
                AppConfig.SpeakervoiceIndexNew++;
            }
            var selectedVoice = voices[AppConfig.SpeakervoiceIndexNew];
            synth.SelectVoice(selectedVoice.VoiceInfo.Name);
        }

        private void StartMailRecognizer()
        {
            if (MailBox.MailList.Count == 0)
                RuntimeConfig.EmptyBox = true;
            if (RuntimeConfig.SettingsMailNew)
            {
                synth.Speak(Statements.MailBoxNewMail(MailBox.NewList.Count));
                synth.Speak(Statements.ReadNew);
                EnableRecognizer(GrammarRepo.YesNo);
            }
            else
            {
                ReadMailItem();
                EnableRecognizer(GrammarRepo.Actions);
            }
        }

        private void LoadGrammars()
        {
            recognizer.LoadGrammarAsync(GrammarRepo.Actions);
            recognizer.LoadGrammarAsync(GrammarRepo.Continue);
            recognizer.LoadGrammarAsync(GrammarRepo.Read);
            recognizer.LoadGrammarAsync(GrammarRepo.Settings);
            recognizer.LoadGrammarAsync(GrammarRepo.SettingsOptions);
            recognizer.LoadGrammarAsync(GrammarRepo.SettingsVoice);
            recognizer.LoadGrammarAsync(GrammarRepo.YesNo);
            recognizer.LoadGrammarAsync(GrammarRepo.Folders);
        }

        #region Recognizers

        private void EnableRecognizer(Grammar grammarName)
        {

            RuntimeConfig.GrammarIndex = recognizer.Grammars.IndexOf(grammarName);
            recognizer.Grammars[RuntimeConfig.GrammarIndex].Enabled = true;
            recognizer.Recognize();

        }

        #endregion

        #region Events
        void Application_NewMailEx(string EntryIDCollection)
        {
            if (AppConfig.MailBoxReadNew)
            {
                RuntimeConfig.MailNew = true;
                MailBox.NewList.Push(items.GetLast());
                if (RuntimeConfig.EmptyBox)
                    StartMailRecognizer();
            }
        }

        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            recognizer.Grammars[RuntimeConfig.GrammarIndex].Enabled = false;
            //recognizer.UnloadAllGrammars();
            timer.Enabled = false;

            var result = e.Result.Text.ToLower();
            if (RuntimeConfig.Move)
                RuntimeConfig.FolderName = result;
            else
            {
                switch (result)
                {
                    case "read folders":
                    case "what are my folder options":
                    case "folder options":
                    case "what folders":
                        RuntimeConfig.ReaderActionsFolderNames = true;
                        break;
                    case "go":
                        RuntimeConfig.Go = true;
                        break;
                    case "yes":
                        RuntimeConfig.Yes = true;
                        RuntimeConfig.No = false;
                        break;
                    case "no":
                        RuntimeConfig.No = true;
                        RuntimeConfig.Yes = false;
                        break;
                    case "remove":
                    case "delete":
                    case "deeleet":
                    case "deeleat":
                        RuntimeConfig.Delete = true;
                        break;
                    case "move":
                    case "move to folder":
                    case "folder":
                    case "folders":
                        RuntimeConfig.Move = true;
                        break;
                    case "skip":
                    case "next":
                    case "necks":
                        RuntimeConfig.Skip = true;
                        break;
                    case "repeat":
                        RuntimeConfig.Repeat = true;
                        break;
                    case "read":
                        RuntimeConfig.Read = true;
                        break;
                    case "options":
                        RuntimeConfig.ReadFolders = true;
                        break;
                    case "settings":
                        RuntimeConfig.Settings = true;
                        break;
                    case "mail":
                        RuntimeConfig.SettingsMailNew = true;
                        break;
                    case "voice":
                        RuntimeConfig.SettingsVoice = true;
                        break;
                    case "change voice":
                    case "new voice":
                        RuntimeConfig.SettingsVoiceChange = true;
                        break;
                    case "talk faster":
                    case "faster":
                        RuntimeConfig.SettingsVoiceRate = true;
                        RuntimeConfig.SettingsVoiceRateIncrease = true;
                        break;
                    case "slower":
                    case "talk slower":
                        RuntimeConfig.SettingsVoiceRate = true;
                        RuntimeConfig.SettingsVoiceRateDecrease = true;
                        break;
                    case "stop":
                        RuntimeConfig.SpeakerStop = true;
                        break;
                }
            }
            Router();
        }

        private void Router()
        {
            if (RuntimeConfig.Start)
            {
                if (RuntimeConfig.Yes)
                {
                    ReadMailItem();
                    RuntimeConfig.Reset();
                    EnableRecognizer(GrammarRepo.Actions);
                }
                else if (RuntimeConfig.No)
                {
                    synth.Speak(Statements.SpeakerPause);
                    synth.Speak(Statements.SpeakerGo);
                    RuntimeConfig.Reset();
                    EnableRecognizer(GrammarRepo.Continue);
                }
            }
            if (RuntimeConfig.ReaderActionsFolderNames)
            {
                RuntimeConfig.ReaderActionsFolderNames = false;
                synth.Speak(Statements.ActionMoveFolderList(MailBox.Folders));
            }
            if (RuntimeConfig.Settings)
            {
                if (RuntimeConfig.SettingsExit)
                {
                    RuntimeConfig.Reset();
                    StartMailRecognizer();
                }
                else if (RuntimeConfig.SettingsVoice)
                {
                    if (RuntimeConfig.SettingsVoiceChange)
                    {
                        if (RuntimeConfig.Yes)
                        {
                            RuntimeConfig.Reset();
                            AppConfig.SpeakerVoiceIndex = AppConfig.SpeakervoiceIndexNew;
                            synth.Speak(Statements.SettingVoiceChangeConfirmation);
                            StartMailRecognizer();
                        }
                        else
                        {
                            VoiceChoice();
                            synth.Speak(Statements.SettingsVoiceChange(synth.Voice.Name));
                            EnableRecognizer(GrammarRepo.YesNo);
                        }
                    }
                    else if (RuntimeConfig.SettingsVoiceRate)
                    {
                        if (RuntimeConfig.SettingsVoiceRateIncrease)
                        {
                            AppConfig.SpeakerSpeechRate = AppConfig.SpeakerSpeechRate + 1;
                            synth.Rate = AppConfig.SpeakerSpeechRate;
                            Statements.SettingsVoiceSpeedConfirmation(Statements.SettingsVoiceSpeedIncreased);
                        }
                        else if (RuntimeConfig.SettingsVoiceRateDecrease)
                        {
                            AppConfig.SpeakerSpeechRate = AppConfig.SpeakerSpeechRate - 1;
                            synth.Rate = AppConfig.SpeakerSpeechRate;
                            Statements.SettingsVoiceSpeedConfirmation(Statements.SettingsVoiceSpeedDecreased);
                        }
                        synth.Speak(Statements.SettingsExit + " or " + Statements.SettingsVoice);
                        EnableRecognizer(GrammarRepo.SettingsVoice);
                    }
                    else
                    {
                        synth.Speak(Statements.SettingsVoice);
                        EnableRecognizer(GrammarRepo.SettingsVoice);
                    }
                }
                else
                {
                    synth.Speak(Statements.Settings);
                    EnableRecognizer(GrammarRepo.SettingsOptions);
                }
            }
            if (RuntimeConfig.Go)
            {
                ReadMailItem();
                RuntimeConfig.Reset();
                EnableRecognizer(GrammarRepo.Actions);
            }
            if (RuntimeConfig.Read)
            {
                RuntimeConfig.Reset();
                var message = MailInfo.Mail.HTMLBody;
                var body = MailInfo.Mail.Body;
                synth.Speak(body);
                synth.Speak(Questions.Proceed);
                EnableRecognizer(GrammarRepo.Actions);
            }
            if (RuntimeConfig.Skip)
            {
                synth.Speak(Statements.Action(SingleStatements.Skipped));
                RuntimeConfig.Reset();
                ReadMailItem();
                EnableRecognizer(GrammarRepo.Actions);
            }
            if (RuntimeConfig.Delete)
            {
                if (RuntimeConfig.Yes)
                {
                    RuntimeConfig.Reset();
                    MailInfo.Mail.Delete();
                    synth.Speak(Statements.Actions(SingleStatements.Delete));
                    ReadMailItem();
                    EnableRecognizer(GrammarRepo.Actions);
                }
                else if (RuntimeConfig.No)
                {
                    synth.Speak(Questions.Proceed);
                    RuntimeConfig.Reset();
                    EnableRecognizer(GrammarRepo.Actions);
                }
                else
                {
                    synth.Speak(Confirmation.Action(SingleStatements.Delete));
                    EnableRecognizer(GrammarRepo.YesNo);
                }

            }
            if (RuntimeConfig.Repeat)
            {
                ReadMailItem();
                RuntimeConfig.Reset();
                EnableRecognizer(GrammarRepo.Actions);
            }
            if (RuntimeConfig.ReadFolders)
            {
                synth.Speak(Statements.FolderNames);
                foreach (var folder in MailBox.Folders)
                {
                    synth.Speak(folder);
                }
                RuntimeConfig.ReadFolders = false;
            }
            if (RuntimeConfig.Move)
            {
                if (RuntimeConfig.FolderName != string.Empty)
                {
                    var destFolder = inBox.Folders[RuntimeConfig.FolderName];
                    MailInfo.Mail.Move(destFolder);
                    synth.Speak(Statements.ActionMoveConfirmation(destFolder.Name));
                    RuntimeConfig.Reset();
                    ReadMailItem();
                }
                else
                {
                    synth.Speak(Questions.Folder);
                    EnableRecognizer(GrammarRepo.Folders);
                }
            }
        }
        #endregion

        private void GetMailItem()
        {
            if (RuntimeConfig.MailNew)
            {
                MailInfo = MailBox.NewList.Pop();
            }
            else
            {
                MailInfo = MailBox.MailList.Pop();
            }
        }

        private void ReadMailItem()
        {

            if (RuntimeConfig.Repeat)
            {
                RuntimeConfig.Repeat = false;
                synth.Speak(Statements.RepeatMessage);
            }
            else if (RuntimeConfig.Start)
            {
                GetMailItem();
                RuntimeConfig.Start = false;
                synth.Speak(Statements.FirstMessage);
            }
            else
            {
                GetMailItem();
                synth.Speak(Statements.NextMessage);
            }
            synth.Speak(Statements.MailBoxMessageFrom(MailInfo.Mail.SenderName));
            synth.Speak(Statements.MailBoxMessageSubject(MailInfo.Mail.Subject.ToString()));
            synth.Speak(Questions.Proceed);
        }

        private Timer TimeDelay(int seconds)
        {
            var t = new Timer(seconds * 1000) { Enabled = true };

            return t;
        }

        public string ToOrdinal(int num)
        {
            switch (num % 100)
            {
                case 11:
                case 12:
                case 13:
                    return num.ToString() + "th";
            }
            switch (num % 10)
            {
                case 1:
                    return num.ToString() + "st";
                case 2:
                    return num.ToString() + "nd";
                case 3:
                    return num.ToString() + "rd";
                default:
                    return num.ToString() + "th";
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            recognizer.Dispose();
            synth.Dispose();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
