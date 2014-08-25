using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;
using System.Text;
using System.Timers;

namespace OutlookAddIn1.Configuration
{
    public static class AppConfig
    {
        public static int SpeakerSpeechRate = -2;
        public static int SpeakerSpeechRateNew = 0;
        public static double TimerLength = 20000;
        public static bool MailBoxReadNew = true;
        public static bool MailBoxUnreadOnly = true;
        public static int SpeakerVoiceIndex = 0;
        public static int SpeakervoiceIndexNew = SpeakerVoiceIndex;
        public static bool SpeakerNoVoiceRecognition = false;
        public static string MicrosoftVoiceRecognizerName = "MS-1033-80-DESK";
    }
}
