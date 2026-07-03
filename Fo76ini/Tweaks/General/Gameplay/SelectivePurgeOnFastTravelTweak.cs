using System;

namespace Fo76ini.Tweaks.General.Gameplay
{
    class SelectivePurgeOnFastTravelTweak : ITweak<bool>, ITweakInfo
    {
        public string Description => "Enables bSelectivePurgeUnusedOnFastTravel under [BackgroundLoad]. This purges unused assets upon fast traveling to reduce memory usage and stutters.";

        public WarnLevel WarnLevel => WarnLevel.Notice;

        public string AffectedFiles => "Fallout76Custom.ini";

        public string AffectedValues => "[BackgroundLoad]bSelectivePurgeUnusedOnFastTravel";

        public bool DefaultValue => false;

        public string Identifier => this.GetType().FullName;

        public bool UIReloadNecessary => false;

        public bool GetValue()
        {
            return IniFiles.GetBool("BackgroundLoad", "bSelectivePurgeUnusedOnFastTravel", DefaultValue);
        }

        public void SetValue(bool value)
        {
            IniFiles.F76Custom.Set("BackgroundLoad", "bSelectivePurgeUnusedOnFastTravel", value);
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
