using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;
using System.Text;
using VoiceEmailOrganizer.Configuration;

namespace VoiceEmailOrganizer.Repository
{
    public class GrammarRepository
    {
        public Grammar YesNo;
        public Grammar Actions;
        public Grammar Continue;
        public Grammar Settings;
        public Grammar Read;
        public Grammar SettingsOptions;
        public Grammar SettingsVoice;
        public Grammar Folders { get; set; }

        public GrammarRepository()
        {
            YesNo = CreateGrammar(Recognizer.YesNo, Recognizer.YesNoGrammar);
            Actions = CreateGrammar(Recognizer.Actions(), Recognizer.ActionGrammar);
            Continue = CreateGrammar(Recognizer.Continue, Recognizer.GoGrammer);
            Settings = CreateGrammar(Recognizer.Settings, Recognizer.SettingsGrammar);
            SettingsOptions = CreateGrammar(Recognizer.SettingsOptionsList(), Recognizer.SettingsOptionsGrammar);
            SettingsVoice = CreateGrammar(Recognizer.SettingsVoice(), Recognizer.SettingsVoiceGrammar);
            Read = CreateGrammar(Recognizer.Read, Recognizer.ReadGrammar);
        }
        public Grammar CreateGrammar(string[] actions, string grammar)
        {
            var mailOptions = new Choices(actions);
            var grammarBuilder = new GrammarBuilder();
            grammarBuilder.Append(mailOptions);

            var grammars = new Grammar(grammarBuilder) { Name = grammar };
            return grammars;
        }
    }
}
